#Get all Mailboxes into temp variable
$Mailboxes = Get-Mailbox -ResultSize Unlimited


#Export 2 export mailboxes’ smtp aliases
$mailboxes | Get-Mailbox -ResultSize Unlimited | Select-Object RecipientTypeDetails,PrimarySmtpAddress -ExpandProperty emailaddresses | select RecipientTypeDetails,PrimarySmtpAddress, @{name="SMTPALIAS";expression={$_}} | Export-Csv C:\temp\dg\mailbox-SMTPproxy.csv -NoTypeInformation