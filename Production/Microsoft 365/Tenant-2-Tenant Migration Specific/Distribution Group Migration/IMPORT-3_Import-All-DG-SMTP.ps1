$SMTPALIAS = Import-Csv C:\temp\dg\distributiongroups-SMTPproxy_modified.csv
$SMTPALIAS | % {Set-DistributionGroup -Identity $_.PrimarySmtpAddress -EmailAddresses @{Add=$_.SMTPALIAS}}
