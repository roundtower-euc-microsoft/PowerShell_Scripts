$SMTPALIAS = Import-Csv C:\temp\dg\mailbox-SMTPproxy.csv
$SMTPALIAS | % {Set-Mailbox -Identity $_.PrimarySmtpAddress -EmailAddresses @{Add=$_.SMTPALIAS}}