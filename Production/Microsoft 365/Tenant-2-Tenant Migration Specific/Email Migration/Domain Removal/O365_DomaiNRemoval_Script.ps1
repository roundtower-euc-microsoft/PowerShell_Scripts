##############################################################
##                                                          ##
#        Office 365 Migration - Domain Removal Script        #
#                                                            #
#                  By: Corey St. Pierre                      #
#                      Sr. Microsoft Systems Engineer        #
#                      Ahead, LLC                            #
#                      corey.stpierre@ahead.com              #
##                                                          ##
##############################################################

<#
    ./SYNOPSIS

        The purpose of this script is to provide a way for
        administrators to remove a domain from an Office 365
        tenant post tenant-2-tenant migration. 

    ./WARNING
        This script is intended to be used as is. It has 
        been thoroughly tested for its intended purposes.
        Any modifications to this script without the 
        creators consent can result in loss of script
        functionality and/or data loss. The creator is 
        not responisble for any data loss due to misuse
        of said script.

    ./SOURCES
    JiJi Technologies - How to remove an Office 365 domain using PowerShell
    https://blog.jijitechnologies.com/how-to-remove-an-office365-domain-using-powershell/

    https://gallery.technet.microsoft.com/office/Office-365-Connection-47e03052

#>

## Self Elevating Permission
## Get the ID and security principal of the current user account
 $myWindowsID=[System.Security.Principal.WindowsIdentity]::GetCurrent()
 $myWindowsPrincipal=new-object System.Security.Principal.WindowsPrincipal($myWindowsID)


 ## Get the security principal for the Administrator role
 $adminRole=[System.Security.Principal.WindowsBuiltInRole]::Administrator


 ## Check to see if we are currently running "as Administrator"
 If ($myWindowsPrincipal.IsInRole($adminRole))
    {
    ## We are running "as Administrator" - so change the title and background color to indicate this
    $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)"
    $Host.UI.RawUI.BackgroundColor = "DarkBlue"
    Clear-Host
    }
 Else
    {
    ## We are not running "as Administrator" - so relaunch as administrator
    
    ## Create a new process object that starts PowerShell
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell";
    
    ## Specify the current script path and name as a parameter
    $newProcess.Arguments = $myInvocation.MyCommand.Definition;
    
    ## Indicate that the process should be elevated
    $newProcess.Verb = "runas";
    
    ## Start the new process
    [System.Diagnostics.Process]::Start($newProcess);
    
    ## Exit from the current, unelevated, process
    Exit
    }

##Checking for MSOnline Module
    $MSOModulePath = ((Get-ChildItem -Path "C:\Program Files\WindowsPowerShell\Modules\MSOnline\" -Filter "MSOnline.psd1" -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)

if ($MSOModulePath -ne $null){
    Write-Host "The Microsoft Online Services PowerShell Module is present, continuing script..." -BackgroundColor Black -ForegroundColor Green
    }
 
else{
    Write-Host "The Microsoft Online Services PowerShell Module is not installed. Installing Now....."
    Install-Module MSOnline -Force
    Write-Host "Microsoft Online Services PowerShell Module is now installed. Continuing....."
    Start-Sleep 1
	}

##Checking for Exchange Online Module
    $EXOModulePath = ((Get-ChildItem -Path "C:\Program Files\WindowsPowerShell\Modules\ExchangeOnlineManagement\" -Filter ExchangeOnlineManagement.psd1 -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)

if ($EXOModulePath -ne $null){
	Write-Host "The Exchange Online PowerShell Module is present, continuing script..." -BackgroundColor Black -ForegroundColor Green
	}
 
else{
	Write-Host "The Microsoft Exchange Online PowerShell Module is not installed. Installing Now....."
	Install-Module ExchangeOnlineManagement -Force
	Write-Host "Exchange Online PowerShell module is now installed. Continuing....."
	Start-Sleep 1
	}

$MenuList = {

    Write-Host "	**********************************************************************" -ForegroundColor White
    Write-Host "	                 Remove Domain from Office 365 Tenant                 " -ForegroundColor Cyan
    Write-Host "	**********************************************************************" -ForegroundColor White
    Write-Host "		"
    Write-Host ''
    Write-Host "	Remove Domain from O365" -ForegroundColor Cyan
    Write-Host "	-----------------------" -ForegroundColor Cyan
    Write-Host "	1) Connect to MSOnline & Exchange Online PowerShell" -ForegroundColor White
    Write-Host "	2) Disable AAD Connect Directory Sync" -ForegroundColor White
    Write-Host "	3) Set current & .onmicrosoft.com domain variable" -ForegroundColor White
    Write-Host "	4) Change the UserPrincipalName for all Office 365 users" -ForegroundColor White
    Write-Host "	5) Change Email Addresses for all Office 365 Mailboxes" -ForegroundColor White
    Write-Host "	6) Change Email Addresses for all Distribution Groups" -ForegroundColor White
    Write-Host "	7) Change Email Addresses for all Office 365 Groups" -ForegroundColor White
    Write-Host "	8) Remove Domain Email Addresses from All Mailboxes" -ForegroundColor White
    Write-Host "	9) Remove Domain Email Addresses from All Groups" -ForegroundColor White
    Write-Host "	10) Remove Domain Email Addresses from All Office 365 Groups" -ForegroundColor White
    Write-Host "	11) Remove Domain from Office 365" -ForegroundColor White
    Write-Host "	12) Check for Problematic Remaining Users" -ForegroundColor White
    Write-Host "	"
    Write-Host "	13) Exit" -ForegroundColor Cyan
    Write-Host "	"
    Write-Host "	Select an option.. [1-13]? " -NoNewLine
}

Do {
	
	If ($Choice -ne "None") {Write-Host "Last command: "$Choice -ForegroundColor Yellow}	
	Invoke-Command -ScriptBlock $MenuList
    $Choice = Read-Host

    switch ($Choice)    {

    1 {#   Connect to MSOnline & Exchange Online PowerShell
        
        Write-host "Connecting to MSOnline, please provide your admin credentials" -ForegroundColor Cyan -BackgroundColor Black
        Connect-MsolService
        Start-Sleep 1
        Write-host "Connecting to Exchange Online, please provide your admin credentials" -ForegroundColor Cyan -BackgroundColor Black
        Connect-ExchangeOnline
        Start-Sleep 1
    }

    2 {#   Disable AAD Connect Directory Sync

        Write-host "Disabling AAD Connect Directory Sync" -ForegroundColor Cyan -BackgroundColor Black
        Set-MsolDirSyncEnabled -EnableDirSync $false
    }

    3 {#   Set current & .onmicrosoft.com domain variable
        
        $olddomain = Read-host "Please Set the Domain Name that will be Removed"
        $Newdomain= Read-host "Please Set the .onmicrosoft.com Domain Name that will be set"
    }

    4 {#   Change the UserPrincipalName for all Office 365 users

        $users=Get-MsolUser -domain $olddomain
        $users | Foreach-Object{ 
        $user=$_
        $UserName =($user.UserPrincipalName -split "@")[0]
        $UPN= $UserName+"@"+ $Newdomain 
        Set-MsolUserPrincipalName -UserPrincipalName $user.UserPrincipalName -NewUserPrincipalName $UPN
        }
    }

    5 {#   Change Primary Email Addresses for all Office 365 Mailboxes

        $Users=Get-Mailbox -ResultSize Unlimited
        $Users | Foreach-Object{ 
        $user=$_
        $UPN=$User.UserPrincipalName
        $UserName =($user.PrimarySmtpAddress -split "@")[0]
        $SMTP ="SMTP:"+ $UserName +"@"+$Newdomain 
        $Emailaddress=$UserName+"@"+$Newdomain
        Set-Mailbox $UPN -EmailAddresses $SMTP -WindowsEmailAddress $Emailaddress
        } 
    }

    6 {#   Change Primary Email Addresses for all Distribution Groups

        $Groups=Get-DistributionGroup -ResultSize Unlimited
        $Groups | Foreach-Object{ 
        $group=$_
        $groupname =($group.PrimarySmtpAddress -split "@")[0]
        $SMTP ="SMTP:"+$groupname+"@"+$Newdomain 
        $Emailaddress=$groupname+"@"+$Newdomain
        $group |Set-DistributionGroup -EmailAddresses $SMTP -WindowsEmailAddress $Emailaddress
        }
    }

    7 {#   Change Primary Email Addresses for all Office 365 Groups

        $Groups=Get-UnifiedGroup -ResultSize Unlimited
        $Groups | Foreach-Object{ 
        $group=$_
        $groupname =($group.PrimarySmtpAddress -split "@")[0]
        $SMTP ="SMTP:"+$groupname+"@"+$Newdomain 
        $Emailaddress=$groupname+"@"+$Newdomain
        $group | Set-UnifiedGroup -EmailAddresses $SMTP
        $group | Set-UnifiedGroup -PrimarySMTPAddress $EmailAddress
        }
    }


    8 {#   Remove Domain Email Addresses from All Mailboxes

       $RemoveSMTPDomain = "smtp:*@$OldDomain"
       $AllMailboxes = Get-Mailbox -ResultSize Unlimited| Where-Object {$_.EmailAddresses -like $RemoveSMTPDomain}
       ForEach ($Mailbox in $AllMailboxes)
       {  
       $AllEmailAddress  = $Mailbox.EmailAddresses -notlike $RemoveSMTPDomain
       $RemovedEmailAddress = $Mailbox.EmailAddresses -like $RemoveSMTPDomain
       $MailboxID = $Mailbox.PrimarySmtpAddress 
       $MailboxID | Set-Mailbox -EmailAddresses $AllEmailAddress
       }
    }

    9 {#    Remove Domain Email Addresses from all Groups

       $RemoveSMTPDomain = "smtp:*@$OldDomain"
       $AllGroups = Get-DistributionGroup -ResultSize Unlimited| Where-Object {$_.EmailAddresses -like $RemoveSMTPDomain}
       ForEach ($Group in $AllGroups)
       {  
       $AllEmailAddress  = $Group.EmailAddresses -notlike $RemoveSMTPDomain
       $RemovedEmailAddress = $Group.EmailAddresses -like $RemoveSMTPDomain
       $GroupID = $Group.PrimarySmtpAddress 
       $GroupID | Set-DistributionGroup -EmailAddresses $AllEmailAddress
       }
    }

    10 {#    Remove Domain Email Addresses from all Office 365 Groups

       $RemoveSMTPDomain = "smtp:*@$OldDomain"
       $AllGroups = Get-UnifiedGroup -ResultSize Unlimited| Where-Object {$_.EmailAddresses -like $RemoveSMTPDomain}
       ForEach ($Group in $AllGroups)
       {  
       $AllEmailAddress  = $Group.EmailAddresses -notlike $RemoveSMTPDomain
       $RemovedEmailAddress = $Group.EmailAddresses -like $RemoveSMTPDomain
       $GroupID = $Group.PrimarySmtpAddress 
       $GroupID | Set-UnifiedGroup -EmailAddresses $AllEmailAddress
       }
    }


    11 {#   Remove Domain from Office 365

        Remove-MsolDomain -DomainName $olddomain -Force
    }

    12 {#   Check for Problematic Remaining Users

       $IssueUsers = Get-MsolUser -DomainName $olddomain
       $IssueUsers | Out-GridView
    }

    13 {#	Exit

        popd
        Write-Host "Exiting..."
    }
  }

 } While ($Choice -ne 13)

        

            
        