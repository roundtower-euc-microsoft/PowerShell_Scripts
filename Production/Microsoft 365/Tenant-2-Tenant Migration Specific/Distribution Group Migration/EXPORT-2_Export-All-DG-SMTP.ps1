#Get all groups into temp variable
$groups = Get-DistributionGroup -ResultSize Unlimited


#Export 2 export distribution groups’ smtp aliases
$groups | Get-DistributionGroup -ResultSize Unlimited | Select-Object RecipientTypeDetails,PrimarySmtpAddress -ExpandProperty emailaddresses | select RecipientTypeDetails,PrimarySmtpAddress, @{name="SMTPALIAS";expression={$_}} | Export-Csv C:\temp\dg\distributiongroups-SMTPproxy.csv -NoTypeInformation
