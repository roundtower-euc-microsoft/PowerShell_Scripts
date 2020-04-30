##############################################################
##                                                          ##
#     Click-Once Enable Remote Mailbox and InPlace Hold      #
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
        efficient way for all user provisioning admins to 
        enable a remote mailbox for a user (while in Exchange
        Online Hybrid mode) and also enable an InPlace Hold
        as well.

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

#>

#Script Starting
Write-Host "Script is Starting" -BackgroundColor Black -ForegroundColor Yellow

#Setting your On-Premises Credential
Write-Host "Prompting for your UPN & Password for On-Premises Exchange" -BackgroundColor Black -ForegroundColor Yellow
Start-Sleep 2
$UserCredential = Get-Credential
Write-Host "On-Premises Credential Set. Continuing..." -BackgroundColor Black -ForegroundColor Green

#Setting your Cloud UPN for Exchange Online
Write-Host "Prompting for your UPN for Exchange Online" -BackgroundColor Black -ForegroundColor Yellow
Start-Sleep 2
$EXOUserCredential = Get-Credential
Write-Host "Exchange Online Credential Set Credential Set. Continuing..." -BackgroundColor Black -ForegroundColor Green

#Adding Visual Basic Assembly for Box Prompts
Add-Type -AssemblyName Microsoft.VisualBasic

####BEGIN MULTI USER LOOP#####
$title = 'Enable User for RemoteMB & IPH'
$msg   = 'Do you want to proceed with enabling a user for a Remote MB and IPH?'

$yes = New-Object Management.Automation.Host.ChoiceDescription '&Yes'
$no  = New-Object Management.Automation.Host.ChoiceDescription '&No'
$options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$default = 1  # $no

do{
    $response = $Host.UI.PromptForChoice($title, $msg, $options, $default)
    if ($response -eq 0) {

		#Asking for SamAccountName of User to be Enabled as Remote Mailbox
		$ReMbox = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the UPN of the user to be enabled as a Remote Mailbox (user@octapharma.com)","$env:ReMbox")
		Write-Host "User Is Set. Continuing..." -BackgroundColor Black -ForegroundColor Green

		#Removing the @octapharma.com from the end of the UPN and Adding the RemoteRoutingAddress
		$SplitVar = $ReMbox
		$SplitVar= $SplitVar -replace '@octapharma.com'
		$HybridProxy = "@octapharmaag.mail.onmicrosoft.com"
		$RemoteRoutingAddress = $SplitVar + $HybridProxy

		#Checking for Microsoft Exchange Online PowerShell Module
		Write-Host "Checking for Presence of EXO PowerShell Module" -BackgroundColor Black -ForegroundColor Yellow


		$EXOModulePath = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)
			if ($EXOModulePath -ne $null){
				Write-Host "The Exchange Online PowerShell Module is present, continuing script..." -BackgroundColor Black -ForegroundColor Green
				}

			else{
				Write-Host "The Microsoft Exchange Online PowerShell Module is not installed. Redirecting to download page...."
				Start-Process http://aka.ms/exopspreview
				Write-Host "Script is ending. Please rerun after installing the Microsoft Exchange Online PowerShell Module."
				Start-Sleep 3
				Exit
				}

		#Now Connecting to Exchange On-Premises
		Write-Host "Now Connecting to Exchange On-Premises..." -ForegroundColor Yellow -BackgroundColor Black


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

		#Enabling Remote Mailbox
		Try
		{
			Enable-RemoteMailbox -Identity $SplitVar -RemoteRoutingAddress $RemoteRoutingAddress
			Write-host "Remote mailbox for $ReMbox is enabled, Continuing..." -BackgroundColor Black -ForegroundColor Green
		}
		Catch
		{
		   Write-Host "Enabling Remote Mailbox has failed. Script will continue, however, the RemoteMailbox is not enabled" -BackgroundColor Black -ForegroundColor Red
		}

		#Exiting PSSession with On-Premises Exchange
		Remove-PSSession -Session $Session
		Write-Host "Exited PSSesion with On-Premises Exchange. Continuing..." -BackgroundColor Black -ForegroundColor Yellow

		#Now moving on to the InPlace HOld in Exchange Online. Importing the Module for EXOP.
		Write-Host "Importing the Exchange Online PS Module and Connecting" -BackgroundColor Black -ForegroundColor Yellow
		Import-Module $EXOModulePath
		Connect-EXOPSSession -Credential $EXOUserCredential

		#Setting IPH Variable
		Write-Host "Setting IPH Variable..." -BackgroundColor Black -ForegroundColor Yellow
		$IPHMBNAME = (Get-Mailbox $ReMbox).UserPrincipalName
		$IPHMBNAME = $IPHMBNAME.Split("@")[0]
		$PolicyName = "$IPHMBNAME" + "_20yrIPH"
		Write-Host "IPH Variables Set, Continuting..." -BackgroundColor Black -ForegroundColor Green
		#Setting the InPlace Hold
		Try
		{
			#Add user to their own In Place Hold
			Write-Host "Enabling $ReMbox for IPH on Policy Name $PolicyName" -BackgroundColor Black -ForegroundColor Yellow
			New-MailboxSearch $PolicyName -SourceMailboxes $ReMbox -InPlaceHoldEnabled $true -ItemHoldPeriod 7300
			
			#Disable IMAP and POP3
			Set-CasMailbox -Identity $ReMbox -PopEnabled $false -ImapEnabled $false
			
			Remove-PSSession $Session
			Write-host "$ReMbox has been placed in the $PolicyName hold and has had POP3/IMAP Disabled. Script is ending..." -BackgroundColor Black -ForegroundColor Green
			Start-Sleep 2
		}
		Catch
		{
			Write-Error "$ReMbox - Could Not set In-pace hold - Error $($Error[0])"
		}
	}
} until ($response -eq 1)
Exit
