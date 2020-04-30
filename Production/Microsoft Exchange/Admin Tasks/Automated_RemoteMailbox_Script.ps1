####################################################################################
####################################################################################
#                                                                                  #
# Name: Automated Remote Mailbox Creation Script for Exchange On-Premises          #
# Author: Corey St. Pierre                                                         #
# Company: RoundTower Technologies, LLC                                            #
# Purpose: Automated Script to Enable New Users 2 days older or Less as Remote     #
#          mailboxes in Exchange for Office 365 hybrid purposes.                   #
# Usage: powershell.exe "whateveryounamedthisscript.ps1"                           #
#                                                                                  #
####################################################################################
####################################################################################

#Setting Parameters
$Error.Clear()
$logfiledate = Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
$logfilename = "RUCS_Log_" + "$logfiledate"
$logfilepath = "C:\temp\RemoteUserScriptLogs\$logfilename.txt"
$RemoteRoutingAddress = "@tenant.mail.onmicrosoft.com"
$PSDefaultParameterValues.Add("*-AD*:Server","ci-svr1.ci.ins")
$DomainController = "DC FQDN HERE"


#Importing all PowerShell Modules Required
ac $logfilepath "Importing ActiveDirectory and Exchange PowerShell Modules"
try{
	Import-Module ActiveDirectory -ErrorAction Stop
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
	Set-ADServerSettings -ViewEntireForest $true -PreferredGlobalCatalog
	}
catch{
	ac $logfilepath "Failed to load Active Directory or Exchange module, aborting $($Error[0])"
	Exit
}

#Getting Users to Enable as Remote Mailboxes and Seting to a Variable
ac $logfilepath "Gathering Users to Enable as Remote Mailboxes...."
try{
	$When = ((Get-Date).AddDays(-2)).Date
	$Sams = Get-ADUser -Server ci-svr1.ci.ins -Filter {whenCreated -ge $When} -Properties whenCreated
	foreach($Sam in $Sams) {
        $User = $Sam.SamAccountName
		ac $logfilepath "The User $User was discovered."
	}
}
catch{
	ac $logfilepath "No New Users to Enable as remote mailbox. Ending Script."
	Exit
}

#Enabling the User as a Remote Mailbox
ac $logfilepath "Enabling Gathered mailboxes as Remote Mailboxes in Exchange"
try {
	foreach ($Sam in $Sams) {
        $User = $Sam.SamAccountName
        $RemoteAddress = -join($User,$RemoteRoutingAddress)
		$CheckUser = Get-ADUser $User -Properties * 
		if ($CheckUser.msExchRecipientTypeDetails -eq 2147483648){
			ac $logfilepath "User $User was discovered, but already enabled as Remote Mailbox."
			}
		if ($CheckUser.msExchRecipientTypeDetails -ne 2147483648){
			Enable-RemoteMailbox -Id $Sam.SamAccountName -RemoteRoutingAddress $RemoteAddress -DomainController $DomainController
			ac $logfilepath "The User $User was enabled as a remote mailbox."
			}
		}
	}
catch{
	foreach ($Sam in $Sams) {
        $User = $Sam.SamAccountName
		ac $logfilepath "Failed to Enable $User as a Remote Mailbox"
		Exit
		}
}

#Ending Script
ac $logfilepath "Script is ending"
exit