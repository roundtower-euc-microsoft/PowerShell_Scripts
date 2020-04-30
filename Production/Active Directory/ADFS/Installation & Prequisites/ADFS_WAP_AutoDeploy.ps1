#############################################
#############################################
#                                           #
#  Automated ADFDS & WAP Deployment Script  #
#                                           #
#############################################
#############################################
#
<#

Script by Corey St. Pierre, Sparkhound, LLC
Verion 1.0


.SYNOPSIS

This script was created and intended for the automation of deployment for ADFS and WAP 3.0 on Server 2012 R2. This script WILL NOT work on
any server lower than 2012 R2. Just put your variables in when asked and let it roll!

#>

param(
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[string] $strFilenameTranscript = $MyInvocation.MyCommand.Name + " " + (hostname)+ " {0:yyyy-MM-dd hh-mmtt}.log" -f (Get-Date),
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$true, Mandatory=$false)] 
	[string] $TargetFolder = "c:\Install",
	# [string] $TargetFolder = $Env:Temp
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[bool] $WasInstalled = $false,
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[bool] $RebootRequired = $false,
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[string] $opt = "None",
	[parameter(ValueFromPipeline=$false, ValueFromPipelineByPropertyName=$false, Mandatory=$false)] 
	[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
)
Start-Transcript -path .\$strFilenameTranscript | Out-Null
$error.clear()
# Detect correct OS here and exit if no match (we intentionally truncate the last character to account for service packs)
if ((Get-WMIObject win32_OperatingSystem).Version -notmatch '6.3.9600'){
	Write-Host "`nThis script requires a version of Windows Server 2012 R2, which this is not. Exiting...`n" -ForegroundColor Red
	Exit
}
Clear-Host
Pushd

[string] $menu = @'

	*******************************************
	     ADFS / WAP 3.0 Installation Script
	*******************************************
	
	Please select an option from the list below.
	
	1) Install Primary ADFS Server
	2) install Secondary ADFS Server
	3) Install Primary WAP Server
	4) Install Secondary WAP Server

	98) Restart the Server
	99) Exit

Select an option.. [1-99]?
'@
Do { 	
if ($RebootRequired -eq $true){Write-Host "`t`t`t`t`t`t`t`t`t`n`t`t`t`tREBOOT REQUIRED!`t`t`t`n`t`t`t`t`t`t`t`t`t`n`t`tDO NOT INSTALL ADFS BEFORE RESTARTING!`t`t`n`t`t`t`t`t`t`t`t`t" -backgroundcolor red -foregroundcolor black}
if ($opt -ne "None") {Write-Host "Last command: "$opt -foregroundcolor Yellow}	
$opt = Read-Host $menu

switch ($opt)    {

1 { # Install Primary ADFS Server

$Start = get-date

### Variable Specificateion
write-host -ForegroundColor Yellow "`n`nNow starting Variable Input for Script."

Add-Type -AssemblyName Microsoft.VisualBasic
$CertificateADFSServiceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Name of The ADFS Service (ie adfs.domain.com)","ADFS Service Name","$env:ADFSServiceName")
$ADFSdisplayName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the ADFS Display Name (ie Your Comany HQ, LTD)","ADFS Display Name","$env:ADFSDisplayName")
$CertificateRemotePath = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Path to the Certificate PFX (UNC) (ie \\Server\Share\cert.pfx)","Certificate Path (UNC)","$env:CertPathUNC")
$PfxPasswordADFS = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the PFX File Password","$env:PFXPassword")
Write-host = "Specify your ADFS Service Account Credentials"
$ADFSuser = Get-Credential

write-host -ForegroundColor Yellow "`n`nVariable Input Complete."

### PFX Import
Write-Host -ForegroundColor Yellow "'n'Now Creating Function to pass through PFX Password to Get Thumbprint"

# create a backup of the original cmdlet
if(Test-Path Function:\Get-PfxCertificate){
    Copy Function:\Get-PfxCertificate Function:\Get-PfxCertificateOriginal
}

# create a new cmdlet with the same name (overwrites the original)
function Get-PfxCertificate {
    [CmdletBinding(DefaultParameterSetName='ByPath')]
    param(
        [Parameter(Position=0, Mandatory=$true, ParameterSetName='ByPath')] [string[]] $filePath,
        [Parameter(Mandatory=$true, ParameterSetName='ByLiteralPath')] [string[]] $literalPath,

        [Parameter(Position=1, ParameterSetName='ByPath')] 
        [Parameter(Position=1, ParameterSetName='ByLiteralPath')] [string] $password,

        [Parameter(Position=2, ParameterSetName='ByPath')]
        [Parameter(Position=2, ParameterSetName='ByLiteralPath')] [string] 
        [ValidateSet('DefaultKeySet','Exportable','MachineKeySet','PersistKeySet','UserKeySet','UserProtected')] $x509KeyStorageFlag = 'DefaultKeySet'
    )

    if($PsCmdlet.ParameterSetName -eq 'ByPath'){
        $literalPath = Resolve-Path $filePath 
    }

    if(!$password){
        # if the password parameter isn't present, just use the original cmdlet
        $cert = Get-PfxCertificateOriginal -literalPath $literalPath
    } else {
        # otherwise use the .NET implementation
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($literalPath, $password, $X509KeyStorageFlag)
    }

    return $cert
}
 
Write-Host -ForegroundColor Yellow "'n'Importing ADFS Certificate"

Import-PfxCertificate –FilePath $CertificateADFSremotePath -CertStoreLocation cert:\localMachine\my -Password $PfxPasswordADFS.Password

Write-Host -ForegroundColor "'n'ADFS Certificate Import Complete"

# ADFS Install
write-host -ForegroundColor Yellow "`n`Now Installing ADFS and Configuring"

Add-WindowsFeature ADFS-Federation -IncludeManagementTools
Import-Module ADFS
$CertificateThumbprint = Get-PfxCertificate -FilePath $CertificateRemotePath $PfxPasswordADFS | Select Thumbprint
Install-AdfsFarm -CertificateThumbprint $CertificateThumbprint -FederationServiceDisplayName $ADFSdisplayName -FederationServiceName $CertificateADFSServiceName -ServiceAccountCredential $ADFSuserCredential
 
write-host -ForegroundColor Yellow "`n`ADFS Installation is Complete."

$Finish = get-date
$Elapsed = $finish - $start
"Elapsed time: {0:mm} minutes and {0:ss} seconds" -f $Elapsed
$RebootRequired = $False
}

2 { # Install Secondary ADFS Server

$Start = get-date

### Variable Specificateion
write-host -ForegroundColor Yellow "`n`nNow starting Variable Input for Script."

Add-Type -AssemblyName Microsoft.VisualBasic
$CertificateADFSServiceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Name of The ADFS Service (ie adfs.domain.com)","ADFS Service Name","$env:ADFSServiceName")
$ADFSdisplayName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the ADFS Display Name (ie Your Comany HQ, LTD)","ADFS Display Name","$env:ADFSDisplayName")
$CertificateRemotePath = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Path to the Certificate PFX (UNC) (ie \\Server\Share\cert.pfx)","Certificate Path (UNC)","$env:CertPathUNC")
$PfxPasswordADFS = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the PFX File Password","$env:PFXPassword")
Write-host = "Specify your ADFS Service Account Credentials"
$ADFSuser = Get-Credential

write-host -ForegroundColor Yellow "`n`nVariable Input Complete."

### PFX Import
Write-Host -ForegroundColor Yellow "'n'Now Creating Function to pass through PFX Password to Get Thumbprint"

# create a backup of the original cmdlet
if(Test-Path Function:\Get-PfxCertificate){
    Copy Function:\Get-PfxCertificate Function:\Get-PfxCertificateOriginal
}

# create a new cmdlet with the same name (overwrites the original)
function Get-PfxCertificate {
    [CmdletBinding(DefaultParameterSetName='ByPath')]
    param(
        [Parameter(Position=0, Mandatory=$true, ParameterSetName='ByPath')] [string[]] $filePath,
        [Parameter(Mandatory=$true, ParameterSetName='ByLiteralPath')] [string[]] $literalPath,

        [Parameter(Position=1, ParameterSetName='ByPath')] 
        [Parameter(Position=1, ParameterSetName='ByLiteralPath')] [string] $password,

        [Parameter(Position=2, ParameterSetName='ByPath')]
        [Parameter(Position=2, ParameterSetName='ByLiteralPath')] [string] 
        [ValidateSet('DefaultKeySet','Exportable','MachineKeySet','PersistKeySet','UserKeySet','UserProtected')] $x509KeyStorageFlag = 'DefaultKeySet'
    )

    if($PsCmdlet.ParameterSetName -eq 'ByPath'){
        $literalPath = Resolve-Path $filePath 
    }

    if(!$password){
        # if the password parameter isn't present, just use the original cmdlet
        $cert = Get-PfxCertificateOriginal -literalPath $literalPath
    } else {
        # otherwise use the .NET implementation
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($literalPath, $password, $X509KeyStorageFlag)
    }

    return $cert
}
 
Write-Host -ForegroundColor Yellow "'n'Importing ADFS Certificate"

Import-PfxCertificate –FilePath $CertificateADFSremotePath -CertStoreLocation cert:\localMachine\my -Password $PfxPasswordADFS.Password

Write-Host -ForegroundColor "'n'ADFS Certificate Import Complete"

# ADFS Install
write-host -ForegroundColor Yellow "`n`Now Installing ADFS and Configuring"

Add-WindowsFeature ADFS-Federation -IncludeManagementTools
Import-Module ADFS
$CertificateThumbprint = Get-PfxCertificate -FilePath $CertificateRemotePath $PfxPasswordADFS | Select Thumbprint
Add-AdfsFarmNode -CertificateThumbprint $CertificateThumbprint -ServiceAccountCredential $ADFSuserCredential -PrimaryComputerName $ADFSprimaryServer -PrimaryComputerPort 80
 
write-host -ForegroundColor Yellow "`n`ADFS Installation is Complete."

$Finish = get-date
$Elapsed = $finish - $start
"Elapsed time: {0:mm} minutes and {0:ss} seconds" -f $Elapsed
$RebootRequired = $False
}

3 { # Install Primary WAP Server

$Start = get-date
 
### Variable Specification
write-host -ForegroundColor Yellow "`n`nNow starting Variable Input for Script."

Add-Type -AssemblyName Microsoft.VisualBasic
$CertificateADFSServiceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Name of The ADFS Service (ie adfs.domain.com)","ADFS Service Name","$env:ADFSServiceName")
$CertificateRemotePath = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Path to the Certificate PFX (UNC) (ie \\Server\Share\cert.pfx)","Certificate Path (UNC)","$env:CertPathUNC")
$PfxPasswordADFS = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the PFX File Password","$env:PFXPassword")
Write-host = "Specify your ADFS Service Account Credentials"
$ADFSuser = Get-Credential

write-host -ForegroundColor Yellow "`n`nVariable Input Complete."

### Verify SNI
write-host -ForegroundColor Yellow "`n`nVerifying SNI is configured properly"

$SniIPport = "0.0.0.0:443" # IP and port to bind to. 0.0.0.0:443 matches all

write-host -ForegroundColor Yellow "`n`nSNI Properly Configured"

### PFX Import
Write-Host -ForegroundColor Yellow "'n'Now Creating Function to pass through PFX Password to Get Thumbprint"

# create a backup of the original cmdlet
if(Test-Path Function:\Get-PfxCertificate){
    Copy Function:\Get-PfxCertificate Function:\Get-PfxCertificateOriginal
}

# create a new cmdlet with the same name (overwrites the original)
function Get-PfxCertificate {
    [CmdletBinding(DefaultParameterSetName='ByPath')]
    param(
        [Parameter(Position=0, Mandatory=$true, ParameterSetName='ByPath')] [string[]] $filePath,
        [Parameter(Mandatory=$true, ParameterSetName='ByLiteralPath')] [string[]] $literalPath,

        [Parameter(Position=1, ParameterSetName='ByPath')] 
        [Parameter(Position=1, ParameterSetName='ByLiteralPath')] [string] $password,

        [Parameter(Position=2, ParameterSetName='ByPath')]
        [Parameter(Position=2, ParameterSetName='ByLiteralPath')] [string] 
        [ValidateSet('DefaultKeySet','Exportable','MachineKeySet','PersistKeySet','UserKeySet','UserProtected')] $x509KeyStorageFlag = 'DefaultKeySet'
    )

    if($PsCmdlet.ParameterSetName -eq 'ByPath'){
        $literalPath = Resolve-Path $filePath 
    }

    if(!$password){
        # if the password parameter isn't present, just use the original cmdlet
        $cert = Get-PfxCertificateOriginal -literalPath $literalPath
    } else {
        # otherwise use the .NET implementation
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($literalPath, $password, $X509KeyStorageFlag)
    }

    return $cert
}
 
Write-Host -ForegroundColor Yellow "'n'Importing ADFS Certificate"

Import-PfxCertificate –FilePath $CertificateADFSremotePath -CertStoreLocation cert:\localMachine\my -Password $PfxPasswordADFS.Password

Write-Host -ForegroundColor "'n'ADFS Certificate Import Complete"

## Add Web Application Proxy Role
Install-WindowsFeature Telnet-Client, RSAT-AD-PowerShell, Web-Application-Proxy -IncludeManagementTools

# Web Application Proxy Configuration Wizard
$CertificateThumbprint = Get-PfxCertificate -FilePath $CertificateRemotePath $PfxPasswordADFS | Select Thumbprint
Install-WebApplicationProxy -CertificateThumbprint $CertificateADFSThumbprint -FederationServiceName $CertificateADFSServiceName -FederationServiceTrustCredential $ADFScredentials

## Web Application Proxy Applications
$CertificateWAPThumbprint = Get-PfxCertificate -FilePath $CertificateRemotePath $PfxPasswordADFS | Select Thumbprint

## Add ADFS
Add-WebApplicationProxyApplication -Name 'ADFS' -ExternalPreAuthentication PassThrough -ExternalUrl "https://$CertificateADFSServiceName/" -BackendServerUrl "https://$CertificateADFSServiceName/" -ExternalCertificateThumbprint $CertificateWAPThumbprint

$Finish = get-date
$Elapsed = $finish - $start
"Elapsed time: {0:mm} minutes and {0:ss} seconds" -f $Elapsed
$RebootRequired = $False
}

4 { # Install Secondary WAP Server

$Start = get-date
 
### Variable Specification
write-host -ForegroundColor Yellow "`n`nNow starting Variable Input for Script."

Add-Type -AssemblyName Microsoft.VisualBasic
$CertificateADFSServiceName = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Name of The ADFS Service (ie adfs.domain.com)","ADFS Service Name","$env:ADFSServiceName")
$CertificateRemotePath = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the Path to the Certificate PFX (UNC) (ie \\Server\Share\cert.pfx)","Certificate Path (UNC)","$env:CertPathUNC")
$PfxPasswordADFS = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the PFX File Password","$env:PFXPassword")
Write-host = "Specify your ADFS Service Account Credentials"
$ADFSuser = Get-Credential

write-host -ForegroundColor Yellow "`n`nVariable Input Complete."

### Verify SNI
write-host -ForegroundColor Yellow "`n`nVerifying SNI is configured properly"

$SniIPport = "0.0.0.0:443" # IP and port to bind to. 0.0.0.0:443 matches all

write-host -ForegroundColor Yellow "`n`nSNI Properly Configured"

### PFX Import
Write-Host -ForegroundColor Yellow "'n'Now Creating Function to pass through PFX Password to Get Thumbprint"

# create a backup of the original cmdlet
if(Test-Path Function:\Get-PfxCertificate){
    Copy Function:\Get-PfxCertificate Function:\Get-PfxCertificateOriginal
}

# create a new cmdlet with the same name (overwrites the original)
function Get-PfxCertificate {
    [CmdletBinding(DefaultParameterSetName='ByPath')]
    param(
        [Parameter(Position=0, Mandatory=$true, ParameterSetName='ByPath')] [string[]] $filePath,
        [Parameter(Mandatory=$true, ParameterSetName='ByLiteralPath')] [string[]] $literalPath,

        [Parameter(Position=1, ParameterSetName='ByPath')] 
        [Parameter(Position=1, ParameterSetName='ByLiteralPath')] [string] $password,

        [Parameter(Position=2, ParameterSetName='ByPath')]
        [Parameter(Position=2, ParameterSetName='ByLiteralPath')] [string] 
        [ValidateSet('DefaultKeySet','Exportable','MachineKeySet','PersistKeySet','UserKeySet','UserProtected')] $x509KeyStorageFlag = 'DefaultKeySet'
    )

    if($PsCmdlet.ParameterSetName -eq 'ByPath'){
        $literalPath = Resolve-Path $filePath 
    }

    if(!$password){
        # if the password parameter isn't present, just use the original cmdlet
        $cert = Get-PfxCertificateOriginal -literalPath $literalPath
    } else {
        # otherwise use the .NET implementation
        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
        $cert.Import($literalPath, $password, $X509KeyStorageFlag)
    }

    return $cert
}
 
Write-Host -ForegroundColor Yellow "'n'Importing ADFS Certificate"

Import-PfxCertificate –FilePath $CertificateADFSremotePath -CertStoreLocation cert:\localMachine\my -Password $PfxPasswordADFS.Password

Write-Host -ForegroundColor "'n'ADFS Certificate Import Complete"

## Add Web Application Proxy Role
Install-WindowsFeature Telnet-Client, RSAT-AD-PowerShell, Web-Application-Proxy -IncludeManagementTools

# Web Application Proxy Configuration Wizard
$CertificateThumbprint = Get-PfxCertificate -FilePath $CertificateRemotePath $PfxPasswordADFS | Select Thumbprint
Install-WebApplicationProxy -CertificateThumbprint $CertificateADFSThumbprint -FederationServiceName $CertificateADFSServiceName -FederationServiceTrustCredential $ADFScredentials

$Finish = get-date
$Elapsed = $finish - $start
"Elapsed time: {0:mm} minutes and {0:ss} seconds" -f $Elapsed
$RebootRequired = $False
}

98 { # Exit and restart
Stop-Transcript
Restart-Computer 
}

99 { # Exit
if (($WasInstalled -eq $false) -and (Get-Module BitsTransfer)){
Write-Host "BitsTransfer: Removing..." -NoNewLine
Remove-Module BitsTransfer
Write-Host "`b`b`b`b`b`b`b`b`b`b`bremoved!   " -ForegroundColor Green
}
popd
Write-Host "Exiting..."
Stop-Transcript
}
default {Write-Host "You haven't selected any of the available options. "}
}
} while ($opt -ne 99)