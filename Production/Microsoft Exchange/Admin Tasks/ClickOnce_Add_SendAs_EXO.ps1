##############################################################
##                                                          ##
#      Click-Once Add SendAs Mailbox Permissions Script      #
#                                                            #
#                  By: Corey St. Pierre                      #
#                      Sr. Microsoft Systems Engineer        #
#                      RoundTower Technologies, LLC          #
#                      corey.stpierre@roundtower.com         #
##                                                          ##
##############################################################

<#
    ./SYNOPSIS

        The purpose of this script is to provide a quick,
        efficient way for help desk and IT Admins to grant
        SendAs mailbox permissions. The script will
        check for the presence of the Exchange Online
        PowerShell module (which supports MFA) and will
        prompt for download and installation in order to
        continue if it is not present. After successful
        verification, prompts for the delegated mailbox's
        alias and the delegate's alias will be presented,
        and then permissions will be granted.

    ./WARNING
        This script is intended to be used as is. It has 
        been thoroughly tested for its intended purposes.
        Any modifications to this script without the 
        creators consent can result in loss of script
        functionality and/or data loss. The creator is 
        not responisble for any data loss due to misuse
        of said script.

    ./SOURCES
    Deva [MSFT] - How to connect to Exchange Online PowerShell using multi-factor authentication??
    https://blogs.msdn.microsoft.com/deva/2019/03/05/how-to-connect-to-exchange-online-powershell-using-multi-factor-authentication/

    https://gallery.technet.microsoft.com/office/Office-365-Connection-47e03052

    Microsoft TechNet
    https://docs.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-permissions-for-recipients

#>

#Script Starting
Write-Host "Script is Starting" -BackgroundColor Black -ForegroundColor Yellow
Write-Host "Prompting for your UPN" -BackgroundColor Black -ForegroundColor Yellow

#Adding Visual Basic Assembly for Box Prompts
Add-Type -AssemblyName Microsoft.VisualBasic
$UPN = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your UPN (i.e. user@octapharma.com","$env:UPN")

#Now Asking for Mailbox that you wish to Delegate SendAs to
#Adding Visual Basic Assembly for Box Prompts
Add-Type -AssemblyName Microsoft.VisualBasic

####BEGIN MULTI USER LOOP#####
$title = 'Grant User SendAs Access'
$msg   = 'Do you want to proceed with granting a user SendAs Access?'

$yes = New-Object Management.Automation.Host.ChoiceDescription '&Yes'
$no  = New-Object Management.Automation.Host.ChoiceDescription '&No'
$options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$default = 1  # $no

do{
    $response = $Host.UI.PromptForChoice($title, $msg, $options, $default)
    if ($response -eq 0) {

	$Delegator = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Email Address of the Mailbox Granting Access (Delegator)","$env:Delegator")
	Write-Host "Delegator is Set." -BackgroundColor Black -ForegroundColor Green

	$Trustee = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Email Address of the Mailbox Being Granted Access (Trustee)","$env:Trustee")
	Write-Host "Trustee is Set." -BackgroundColor Black -ForegroundColor Green

	#Now Connecting to Exchange On-Premises
	Write-Host "Now Connecting to Exchange On-Premises to Add AD Permissions" -ForegroundColor Yellow -BackgroundColor Black

	#Trying First Server
	$ServerTest1 = (Test-Connection -ComputerName se1srv0064.se1.octapharma.net -Quiet)
		If ($ServerTest1 -eq $True)
			{
			$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://se1srv0064.se1.octapharma.net/PowerShell/ -Authentication Kerberos -Credential $UserCredential
			Import-PSSession $Session -DisableNameChecking
			Set-ADServerSettings -ViewEntireForest $true
			}
		#If First Server is not avialable, try second server
		ElseIf ($ServerTest1 -eq $False)
			{
			$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://se1srv0065.se1.octapharma.net/PowerShell/ -Authentication Kerberos -Credential $UserCredential
			Import-PSSession $Session -DisableNameChecking
			Set-ADServerSettings -ViewEntireForest $true
			}
		#If No Servers respond, end script
		else
			{Write-Host "On-Premises Servers are not available, ending script" -BackgroundColor Black -ForegroundColor Red
			Exit
			}
			
	#Adding AD Extended Rights On-Premises
	Write-Host "Now Granting $Trustee SendAs Permissions to $Delegator's Mailbox - On-Premises" -BackgroundColor Black -ForegroundColor Yellow
	Try
	{
		$MbName = (Get-Recipient $Delegator).Name
		Add-ADPermission -Identity "$MbName" -User $Trustee -AccessRights ExtendedRight -ExtendedRights "Send As"
		Write-Host "$Trustee has been granted access successfully on-premises. Continuing...." -BackgroundColor Black -ForegroundColor Green
	}
	Catch
	{
	   Write-Host "$Trustee could not be granted access to $Delegator. Script will continue, but the on-premises process will need to be manually run again..." -BackgroundColor Black -ForegroundColor Red
	}

	#Exiting PSSession with On-Premises Exchange
	Remove-PSSession -Session $Session
	Write-Host "Exited PSSesion with On-Premises Exchange. Continuing..." -BackgroundColor Black -ForegroundColor Yellow

	#Checking for Microsoft Exchange Online PowerShell Module
	Write-Host "Checking for Presence of EXO PowerShell Module" -BackgroundColor Black -ForegroundColor Yellow

	$EXOModulePath = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)
		if ($EXOModulePath -ne $null){
			Write-Host "Importing the Exchange Online PS Module and Connecting" -BackgroundColor Black -ForegroundColor Yellow
			Import-Module $EXOModulePath
			Connect-EXOPSSession -UserPrincipalName $UPN
			}

		else{
			Write-Host "The Microsoft Exchange Online PowerShell Module is not installed. Redirecting to download page...."
			Start-Process http://aka.ms/exopspreview
			Write-Host "Script is ending. Please rerun after installing the Microsoft Exchange Online PowerShell Module."
			Start-Sleep 3
			Exit
			}

	Write-Host "Now Granting $Trustee SendAs Permissions to $Delegator's Mailbox in Exchange Online" -BackgroundColor Black -ForegroundColor Yellow
	Add-RecipientPermission -Identity $Delegator -AccessRights SendAs -Trustee $Trustee -Confirm:$false
	Write-Host "$Trustee has been granted access successfully in Exchange Online. Exiting Script..." -BackgroundColor Black -ForegroundColor Green
	Start-Sleep 3
	}
} until ($response -eq 1)
Exit