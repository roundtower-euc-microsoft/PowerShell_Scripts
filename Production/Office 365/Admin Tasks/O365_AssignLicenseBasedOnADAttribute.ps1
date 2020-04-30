<#
.SYNOPSIS
    The script automatically assigns licenses in Office 365 based on AD-attributes of your choice.
.DESCRIPTION
    Fetches your Office 365 tenant Account SKU's in order to assign licenses based on them.
    Sets mandatory attribute UsageLocation for the user in Office 365 based on local AD attribute Country (or default)
    If switch -MasterOnPremise has been used:
    Assigns Office 365 licenses to the user based on an AD-attribute and reports back to the same attribute if a license was successfully assigned.
    If not:
    Looks after unlicensed users in Office365, checks if the unlicensed user has the right attribute in ad to be verified, otherwise not.
.PARAMETER MasterOnPremise
    Switch, If used, it  Assigns Office 365 licenses to the user based on an AD-attribute and reports back to the same attribute if a license was successfully assigned.    
.PARAMETER Licenses
    Array, defines the licenses used and to activate, specific set of licenses supported. "K1","K2","E1","E3","A1S","A2S","A3S","A1F","A2F","A3F"
.PARAMETER LicenseAttribute
    String, the attribute  used on premise to identify licenses
.PARAMETER AdminUser
    Adminuser in tenant
.PARAMETER DefaultCountry
    Defaultcountry for users if not defined in active directory             
.NOTES
    File Name: ActivateMSOLUsers.ps1
    Author   : Johan Dahlbom, johan[at]dahlbom.eu
    The script are provided “AS IS” with no guarantees, no warranties, and they confer no rights.    
#>
param(
[parameter(Mandatory=$false)]
[Switch]
$MasterOnPremise = $false,

[parameter(Mandatory=$false,HelpMessage="Please provide the SKU's you want to activate")]
[ValidateSet("K1","K2","E1","E3","A1S","A2S","A3S","A1F","A2F","A3F")] 
[ValidateNotNullOrEmpty()] 
[array]
$Licenses,

[parameter(Mandatory=$false)]
[string]
$LicenseAttribute = "employeeType",

[parameter(Mandatory=$false)]
[string]
$AdminUser = "admin@tenant.onmicrosoft.com",

[parameter(Mandatory=$false)]
[string]
$DefaultCountry = "SE"
)
#Define variables
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
$Logfile = ($PSScriptRoot + "\MSOL-UserActivation.log")    
$TimeDate = Get-Date -Format "yyyy-MM-dd-HH:mm"
$SupportedSKUs = @{
                    "K1" = "DESKLESSPACK"
                    "K2" = "DESKLESSWOFFPACK"
                    "E1" = "STANDARDPACK"		
                    "E3" = "ENTERPRISEPACK"
                    "A1S" = "STANDARDPACK_STUDENT"
                    "A2S" = "STANDARDWOFFPACKPACK_STUDENT"
                    "A3S" = "ENTERPRISEPACK_STUDENT"
                    "A1F" = "STANDARDPACK_FACULTY"
                    "A2F" = "STANDARDWOFFPACKPACK_FACULTY"
                    "A3F" = "ENTERPRISEPACK_FACULTY"
                    
                    }
    
################################################
##### Functions

#Function to Import Modules with error handling
Function Import-MyModule 
{ 
Param([string]$name) 
    if(-not(Get-Module -name $name)) 
    { 
        if(Get-Module -ListAvailable $name) 
        { 
            Import-Module -Name $name 
            Write-Host "Imported module $name" -ForegroundColor Yellow
        } 
        else 
        { 
            Throw "Module $name is not installed. Exiting..." 
        }
    } 
    else 
    { 
        Write-Host "Module $name is already loaded..." -ForegroundColor Yellow } 
    }  

#Function to log to file and write output to host
Function LogWrite
{
Param ([string]$Logstring)
Add-Content $Logfile -value $logstring -ErrorAction Stop
Write-Host $logstring 
}
#Function to activate your users based on ad attribute
Function Activate-MsolUsers
{ 
Param([string]$SKU) 

    begin {
        #Set counter to 0
        $i = 0
    }
        
        process {

            #Catch and log errors
            trap {
                $ErrorText = $_.ToString()
		        LogWrite "Error: $_"
                $error.Clear()
                Continue
            }

            #If the switch -MasterOnPremise has been used, start processing licenses from the AD-attribute
	        
            if ($MasterOnPremise) {
        
		    $UsersToActivate = Get-ADUser -filter {$LicenseAttribute -eq $SKU} -Properties $LicenseAttribute -ErrorAction Stop

                if ($UsersToActivate)
			    { 
			    $NumUsers = ($UsersToActivate | Measure-Object).Count

			    LogWrite "Info: $NumUsers user(s) to activate with $SKU"
                foreach($user in $UsersToActivate) { 
                          
                          trap {
                                $ErrorText = $_.ToString()
		                        LogWrite "Error: $_"
                                $error.Clear()
                                Continue
                            }
                        $UPN = $user.userprincipalname
                        $Country = $user.Country
                        LogWrite "Info: Trying to assign licenses to: $UPN"       
                                if (!($country)) { 
                                    $Country = $DefaultCountry } 
                  
                        if ((Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop)) {
	                                 
                                Set-MsolUser -UserPrincipalName $UPN -UsageLocation $country -Erroraction Stop
                                Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $SKUID -Erroraction Stop
									#Verify License Assignment
									if (Get-MsolUser -UserPrincipalName $UPN | Where-Object {$_.IsLicensed -eq $true}) {
										Set-ADUser $user -Replace @{$LicenseAttribute=$SKU+' - Licensed at ' + $TimeDate}
										LogWrite "Info: $upn successfully licensed with $SKU"
                                        $i++;
									} 
									else
									{
										LogWrite "Error: Failed to license $UPN with $SKU, please do further troubleshooting"

									}
                        }
				    }
			    }
            }
	#If no switch has been used, process users and licenses from MSOnline
            else {
			    $UsersToActivate = Get-MsolUser -UnlicensedUsersOnly -All
				    if ($Userstoactivate)				
				    { 
				    $NumUsers = $UsersToActivate.Count
				    LogWrite "Info: $NumUsers unlicensed user(s) in tenant: $($SKUID.ToLower().Split(':')[0]).onmicrosoft.com"
				    foreach ($user in $UsersToActivate) {
                                trap {
                                        $ErrorText = $_.ToString()
		                                LogWrite "Error: $_"
                                        $error.Clear()
                                        Continue
                                }

					        $UPN = $user.UserPrincipalName
					        $ADUser = Get-Aduser -Filter {userPrincipalName -eq $UPN} -Properties $LicenseAttribute -ErrorAction Stop
					        $Country = $ADUser.Country
							    if (!($Country)) { 
								$Country = $DefaultCountry 
							    } 
					        if ($ADUser.$LicenseAttribute -eq $SKU) {
						        LogWrite "Info: Trying to assign licenses to: $UPN"   
						        Set-MsolUser -UserPrincipalName $UPN -UsageLocation $country -Erroraction Stop
						        Set-MsolUserLicense -UserPrincipalName $UPN -AddLicenses $SKUID -Erroraction Stop
					
						        #Verify License Assignment
						        if (Get-MsolUser -UserPrincipalName $UPN | Where-Object {$_.IsLicensed -eq $true}) {
							    LogWrite "Info: $upn successfully licensed with $SKU" 	
                                $i++;
						        } 
						   else
						        {
								
                                LogWrite "Error: Failed to license $UPN with $SKU, please do further troubleshooting"

						        }
				        }
			        }		
		        }
	        }
        }

    End{
    LogWrite "Info: $i user(s) was activated with $license ($SKUID)"
	}
}
################################################
#Import modules required for the script to run
Import-MyModule MsOnline
Import-MyModule ActiveDirectory

	#Start logging and check logfile access
	try 
	{
		LogWrite -Logstring "**************************************************`r`nLicense activation job started at $timedate`r`n**************************************************"
	} 
	catch 
	{
		Throw "You don't have write permissions to $logfile, please start an elevated PowerShell prompt or change NTFS permissions"
	}
	
    #Connect to Azure Active Directory
	try 
	{	 
		$PasswordFile = ($PSScriptRoot + "\$adminuser.txt")
		if (!(Test-Path  -Path $passwordfile))
			{
				Write-Host "You don't have an admin password assigned for $adminuser, please provide the password followed with enter below:" -Foregroundcolor Yellow
				Read-Host -assecurestring | convertfrom-securestring | out-file $passwordfile
			}
		$password = get-content $passwordfile | convertto-securestring -ErrorAction Stop
		$credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $adminuser,$password -ErrorAction Stop
		Connect-MsolService -credential $credentials -ErrorAction Stop
	} 
	catch [system.exception]
	{
        $err = $_.Exception.Message
		LogWrite "Error: Could not Connect to Office 365, please check connection, username and password.`r`nError: $err"
        exit
	}

if (!$licenses) {
    LogWrite "Error: No licenses specified, please specify a supported license`r`nInfo: Supported licenses are: K1,K2,E1,E3,A1S,A2S,A3S,A1F,A2F,A3F!"
}
else
{

#Start processing license assignment
    foreach($license in $Licenses) {

        $SKUID = (Get-MsolAccountSku | Where-Object {$_.AccountSkuId -like "*:$($SupportedSKUs.Get_Item($license))"}).AccountSkuId
	
            if ($SKUID) 
	        {
		        Activate-MsolUsers -SKU $license 
	        }
	        else
	        {
		    LogWrite "Error: No $license licenses in your tenant!"
	        }
    }
}
LogWrite -Logstring "**************************************************"