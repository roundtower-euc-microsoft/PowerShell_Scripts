##############################################################
##                                                          ##
#   Click-Once Add Full Access Mailbox Permissions Script    #
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
        Full Access mailbox permissions. The script will
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
    https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/add-mailboxpermission?view=exchange-ps

#>

#Script Starting
Write-Host "Script is Starting" -BackgroundColor Black -ForegroundColor Yellow
Write-Host "Prompting for your UPN" -BackgroundColor Black -ForegroundColor Yellow

#Adding Visual Basic Assembly for Box Prompts
Add-Type -AssemblyName Microsoft.VisualBasic
$UPN = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your UPN (i.e. user@octapharma.com","$env:UPN")

#Checking for Microsoft Exchange Online PowerShell Module
Write-Host "Checking for Presence of EXO PowerShell Module" -BackgroundColor Black -ForegroundColor Yellow

$EXOModulePath = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)
    if ($EXOModulePath -ne $null){
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

#Now Asking for Mailbox that you wish to Delegate Full Access to
#Adding Visual Basic Assembly for Box Prompts
Add-Type -AssemblyName Microsoft.VisualBasic

####BEGIN MULTI USER LOOP#####
$title = 'Add FullAccess Rights'
$msg   = 'Do you want to proceed with granting a user FullAccess to another mailbox?'

$yes = New-Object Management.Automation.Host.ChoiceDescription '&Yes'
$no  = New-Object Management.Automation.Host.ChoiceDescription '&No'
$options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$default = 1  # $no

do{
		$response = $Host.UI.PromptForChoice($title, $msg, $options, $default)
		if ($response -eq 0) {

	$Delegator = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Email Address of the Mailbox Granting Access (Delegator)","$env:Delegator")
	Write-Host "Delegator is Set." -BackgroundColor Black -ForegroundColor Green

	$Delegate = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Email Address of the Mailbox Being Granted Access (Delegate)","$env:Delegate")
	Write-Host "Delegate is Set." -BackgroundColor Black -ForegroundColor Green

	[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

	$caption = "AutoMapping"
	$message = "Does the Delegate require the mailbox be AutoMapped?"
	$yesNoButtons = 4

	if ([System.Windows.Forms.MessageBox]::Show($message, $caption, $yesNoButtons) -eq "NO") {
		"You answered no"

		Write-Host "Now" -BackgroundColor Black -ForegroundColor Yellow
		Add-MailboxPermission -Identity $Delegator -User $Delegate -AccessRights FullAccess -InheritanceType All -AutoMapping $false
		Write-Host "$Delegate has been granted access successfully. Exiting Script..." -BackgroundColor Black -ForegroundColor Green
		Start-Sleep 3
		Exit
	}
	else {
		"You answered yes"
		Write-Host "Now Granting $Delegate Full Access Permissions to $Delegator's Mailbox With AutoMapping" -BackgroundColor Black -ForegroundColor Yellow
		Add-MailboxPermission -Identity $Delegator -User $Delegate -AccessRights FullAccess -InheritanceType All -AutoMapping $true
		Write-Host "$Delegate has been granted access successfully. Exiting Script..." -BackgroundColor Black -ForegroundColor Green
		Start-Sleep 3
	}
}
} until ($response -eq 1)
Exit