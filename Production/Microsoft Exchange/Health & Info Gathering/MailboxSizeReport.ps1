<#

.Requires -version 2 - Runs in Exchange Management Shell

.SYNOPSIS
.\MailboxSizeReport.ps1 - It Can Display all the Mailbox Size with Item Count,Database,Server Details

Or It can Export to a CSV file

Or You can Enter WildCard to Display or Export


Example 1

[PS] C:\>.\MailboxSizeReport.ps1


Mailbox Size Report
----------------------------

1.Display in Exchange Management Shell

2.Export to CSV File

3.Enter the Mailbox Name with Wild Card (Export)

4.Enter the Mailbox Name with Wild Card (Display)

5.Export to CSV File (OFFICE 365)

6.Enter the Mailbox Name with Wild Card (Export) (OFFICE 365)

Choose The Task: 1

Display Name                  Primary SMTP address          TotalItemSize                 ItemCount
------------                  --------------------          -------------                 ---------
Tes433                        Tes433@Welcome.com
Test                          Test@testcareexchange.biz     335.9 KB (343,933 bytes)      40
Test X500                     TestX500@Testexchange.biz     6.544 KB (6,701 bytes)        3
Test100                       test100@testcareexchange.biz  40.74 KB (41,719 bytes)       7
Test22                        Test22@Testexchange.biz       60.04 KB (61,483 bytes)       7
Test3                         Test3@testcareexchange.biz    364.7 KB (373,503 bytes)      31
Test33                        Test332@testcareexchange.biz  93.34 KB (95,585 bytes)       6
Test33                        Test33@FSD.com                5.335 KB (5,463 bytes)        3
Test3331                      Test3331@Testexchange.biz     24.14 KB (24,720 bytes)       2
Test46                        Test46@testcareexchange.biz   254 KB (260,071 bytes)        21

Example 2

[PS] C:\>.\MailboxSizeReport.ps1


Mailbox Size Report
----------------------------

1.Display in Exchange Management Shell

2.Export to CSV File

3.Enter the Mailbox Name with Wild Card (Export)

4.Enter the Mailbox Name with Wild Card (Display)

5.Export to CSV File (OFFICE 365)

6.Enter the Mailbox Name with Wild Card (Export) (OFFICE 365)

Choose The Task: 2
Enter the Path of CSV file (Eg. C:\Report.csv): C:\MailboxReport.csv

.Author
Written By: Satheshwaran Manoharan

Change Log
V1.0, 10/08/2014 - Initial version

Change Log
V1.1, 05/12/2016 - ProgressBar,Seperate Office 365 Options, QuotaLimits,EmailAddresses

#>

Write-host "

Mailbox Size Report
----------------------------

1.Display in Exchange Management Shell

2.Export to CSV File

3.Export to CSV File (Specific to Database)

4.Enter the Mailbox Name with Wild Card (Export)

5.Enter the Mailbox Name with Wild Card (Display)

6.Export to CSV File (OFFICE 365)

7.Enter the Mailbox Name with Wild Card (Export) (OFFICE 365)"-ForeGround "Cyan"

#----------------
# Script
#----------------

Write-Host "               "

$number = Read-Host "Choose The Task"
$output = @()
switch ($number) 
{

1 {

$AllMailbox = Get-mailbox -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount

Write-Output $Userobj

}

;Break}

2 {
$i = 0 

$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\Report.csv)" 

$AllMailbox = Get-mailbox -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue

if (($Mbx.UseDatabaseQuotaDefaults -eq $true) -and (Get-MailboxDatabase $mbx.Database).ProhibitSendReceiveQuota.value -eq $null)
{
$ProhibitSendReceiveQuota = "Unlimited"
}
if (($Mbx.UseDatabaseQuotaDefaults -eq $true) -and (Get-MailboxDatabase $mbx.Database).ProhibitSendReceiveQuota.value -ne $null)
{
(Get-MailboxDatabase $mbx.Database).ProhibitSendReceiveQuota.Value.ToMB()
}
if (($Mbx.UseDatabaseQuotaDefaults -eq $false) -and ($mbx.ProhibitSendReceiveQuota.value -eq $null))
{
$ProhibitSendReceiveQuota = "Unlimited"
}
if (($Mbx.UseDatabaseQuotaDefaults -eq $false) -and ($mbx.ProhibitSendReceiveQuota.value -ne $null))
{
$ProhibitSendReceiveQuota = $Mbx.ProhibitSendReceiveQuota.Value.ToMB()
}

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientTypeDetails
$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.EmailAddresses.smtpaddress -join ";")
$userObj | Add-Member NoteProperty -Name "Database" -Value $mbx.Database
$userObj | Add-Member NoteProperty -Name "ServerName" -Value $mbx.ServerName
if($Stats)
{
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize.Value.ToMB()
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
$userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
$userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize.Value.ToMB()
}
$userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $ProhibitSendReceiveQuota
$userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
$userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime

$output += $UserObj  
# Update Counters and Write Progress
$i++
Write-Progress -Activity "Scanning Mailboxes . . ." -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i/$AllMailbox.Count*100)
}


$output | Export-csv -Path $CSVfile -NoTypeInformation

;Break}

3 {
$i = 0 

$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\Report.csv)" 
$Database = Read-Host "Enter the DatabaseName (Eg. Database 01)" 

$AllMailbox = Get-mailbox -resultsize unlimited -Database "$Database"

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue

if (($Mbx.UseDatabaseQuotaDefaults -eq $true) -and (Get-MailboxDatabase $mbx.Database).ProhibitSendReceiveQuota.value -eq $null)
{
$ProhibitSendReceiveQuota = "Unlimited"
}
if (($Mbx.UseDatabaseQuotaDefaults -eq $true) -and (Get-MailboxDatabase $mbx.Database).ProhibitSendReceiveQuota.value -ne $null)
{
(Get-MailboxDatabase $mbx.Database).ProhibitSendReceiveQuota.Value.ToMB()
}
if (($Mbx.UseDatabaseQuotaDefaults -eq $false) -and ($mbx.ProhibitSendReceiveQuota.value -eq $null))
{
$ProhibitSendReceiveQuota = "Unlimited"
}
if (($Mbx.UseDatabaseQuotaDefaults -eq $false) -and ($mbx.ProhibitSendReceiveQuota.value -ne $null))
{
$ProhibitSendReceiveQuota = $Mbx.ProhibitSendReceiveQuota.Value.ToMB()
}

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientTypeDetails
$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.EmailAddresses.smtpaddress -join ";")
$userObj | Add-Member NoteProperty -Name "Database" -Value $mbx.Database
$userObj | Add-Member NoteProperty -Name "ServerName" -Value $mbx.ServerName
if($Stats)
{
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize.Value.ToMB()
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
$userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
$userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize.Value.ToMB()
}
$userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $ProhibitSendReceiveQuota
$userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
$userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime

$output += $UserObj  
# Update Counters and Write Progress
$i++
Write-Progress -Activity "Scanning Mailboxes . . ." -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i/$AllMailbox.Count*100)
}


$output | Export-csv -Path $CSVfile -NoTypeInformation

;Break}

4 {
$i = 0 
$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\DG.csv)" 

$MailboxName = Read-Host "Enter the Mailbox name or Range (Eg. Mailboxname , Mi*,*Mik)"

$AllMailbox = Get-mailbox $MailboxName -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue

if (($Mbx.UseDatabaseQuotaDefaults -eq $true) -and (Get-MailboxDatabase $mbx.Database).ProhibitSendReceiveQuota.value -eq $null)
{
$ProhibitSendReceiveQuota = "Unlimited"
}
if (($Mbx.UseDatabaseQuotaDefaults -eq $true) -and (Get-MailboxDatabase $mbx.Database).ProhibitSendReceiveQuota.value -ne $null)
{
(Get-MailboxDatabase $mbx.Database).ProhibitSendReceiveQuota.Value.ToMB()
}
if (($Mbx.UseDatabaseQuotaDefaults -eq $false) -and ($mbx.ProhibitSendReceiveQuota.value -eq $null))
{
$ProhibitSendReceiveQuota = "Unlimited"
}
if (($Mbx.UseDatabaseQuotaDefaults -eq $false) -and ($mbx.ProhibitSendReceiveQuota.value -ne $null))
{
$ProhibitSendReceiveQuota = $Mbx.ProhibitSendReceiveQuota.Value.ToMB()
}


$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientTypeDetails
$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.EmailAddresses.smtpaddress -join ";")
$userObj | Add-Member NoteProperty -Name "Database" -Value $mbx.Database
$userObj | Add-Member NoteProperty -Name "ServerName" -Value $mbx.ServerName
if($Stats)
{
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize.Value.ToMB()
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
$userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
$userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize.Value.ToMB()
}
$userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $ProhibitSendReceiveQuota
$userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
$userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime

$output += $UserObj  
# Update Counters and Write Progress
$i++
Write-Progress -Activity "Scanning Mailboxes . . ." -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i/$AllMailbox.Count*100)
}

$output | Export-csv -Path $CSVfile -NoTypeInformation

;Break}

5 {

$MailboxName = Read-Host "Enter the Mailbox name or Range (Eg. Mailboxname , Mi*,*Mik)"

$AllMailbox = Get-mailbox $MailboxName -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount

Write-Output $Userobj

}

;Break}

6 {
$i = 0 
$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\Report.csv)" 

$AllMailbox = Get-mailbox -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientTypeDetails
$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.EmailAddresses -join ";")
$userObj | Add-Member NoteProperty -Name "Database" -Value $Stats.Database
$userObj | Add-Member NoteProperty -Name "ServerName" -Value $Stats.ServerName
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
$userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
$userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize
$userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $Mbx.ProhibitSendReceiveQuota
$userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
$userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime

$output += $UserObj  
# Update Counters and Write Progress
$i++
Write-Progress -Activity "Scanning Mailboxes . . ." -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i/$AllMailbox.Count*100)
}

$output | Export-csv -Path $CSVfile -NoTypeInformation

;Break}

7 {
$i = 0 
$CSVfile = Read-Host "Enter the Path of CSV file (Eg. C:\DG.csv)" 

$MailboxName = Read-Host "Enter the Mailbox name or Range (Eg. Mailboxname , Mi*,*Mik)"

$AllMailbox = Get-mailbox $MailboxName -resultsize unlimited

Foreach($Mbx in $AllMailbox)

{

$Stats = Get-mailboxStatistics -Identity $Mbx.distinguishedname -WarningAction SilentlyContinue

$userObj = New-Object PSObject

$userObj | Add-Member NoteProperty -Name "Display Name" -Value $mbx.displayname
$userObj | Add-Member NoteProperty -Name "Alias" -Value $Mbx.Alias
$userObj | Add-Member NoteProperty -Name "RecipientType" -Value $Mbx.RecipientTypeDetails
$userObj | Add-Member NoteProperty -Name "Recipient OU" -Value $Mbx.OrganizationalUnit
$userObj | Add-Member NoteProperty -Name "Primary SMTP address" -Value $Mbx.PrimarySmtpAddress
$userObj | Add-Member NoteProperty -Name "Email Addresses" -Value ($Mbx.EmailAddresses -join ";")
$userObj | Add-Member NoteProperty -Name "Database" -Value $Stats.Database
$userObj | Add-Member NoteProperty -Name "ServerName" -Value $Stats.ServerName
$userObj | Add-Member NoteProperty -Name "TotalItemSize" -Value $Stats.TotalItemSize
$userObj | Add-Member NoteProperty -Name "ItemCount" -Value $Stats.ItemCount
$userObj | Add-Member NoteProperty -Name "DeletedItemCount" -Value $Stats.DeletedItemCount
$userObj | Add-Member NoteProperty -Name "TotalDeletedItemSize" -Value $Stats.TotalDeletedItemSize
$userObj | Add-Member NoteProperty -Name "ProhibitSendReceiveQuota-In-MB" -Value $Mbx.ProhibitSendReceiveQuota
$userObj | Add-Member NoteProperty -Name "UseDatabaseQuotaDefaults" -Value $Mbx.UseDatabaseQuotaDefaults
$userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $Stats.LastLogonTime

$output += $UserObj  
# Update Counters and Write Progress
$i++
Write-Progress -Activity "Scanning Mailboxes . . ." -Status "Scanned: $i of $($AllMailbox.Count)" -PercentComplete ($i/$AllMailbox.Count*100)
}

$output | Export-csv -Path $CSVfile -NoTypeInformation

;Break}

Default {Write-Host "No matches found , Enter Options 1 or 2" -ForeGround "red"}

}