#Get List of All Mailboxes with Forwarding Addresses
#Created by Corey St. Pierre
#
#
#Step 1 - Connect to Exchange Online through the Windows Azure Active Directory PowerShell Module
#You will be asked to put your Exchange Online Administrative Username and Password in for security
#
#
Import-Module MSOnline
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
#
#
#Step 2 - Get Mailbox Forwarding Information, Format list, and export to CSV
#
#
mkdir C:\MailFowardingListReults\
Get-Mailbox | Where {$_.ForwardingAddress -ne $null} | Select Name, ForwardingAddress, DeliverToMailboxAndForward | Export-CSV C:\MailFowardingListReults\mailforwards.csv