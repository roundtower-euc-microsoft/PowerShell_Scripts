#Mass SMTP Proxy AddressChange
#Written by Corey St. Pierre
#
#Step 1 - Import AD PowerShell Module
Import-Module ActiveDirectory
#
#Step 2 - Set SMTP Proxy Domain Name
$newproxy = Read-Host 'Enter Your SMTP Domain for the Proxy Address'
#
#Step 3 - Set Distinguished Name for Location of Users
$userou = Read-Host 'Specify the Distinguished Name of the OU you want to change (i.e dc=domain,dc=xxx)'
#
#Specify User Search Base
$users = Get-ADUser -Filter '*' -SearchBase $userou -Properties SamAccountName, ProxyAddresses 
#
#Put it all together
Foreach ($user in $users) {Set-ADUser -Identity $user.samaccountname -Add @{Proxyaddresses="SMTP:"+$user.samaccountname+"@"+$newproxy}} 