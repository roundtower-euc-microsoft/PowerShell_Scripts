#Get all groups into temp variable
$groups = Get-DistributionGroup -ResultSize Unlimited


#Export 3 export all distribution groups and members (and member type)
$groups |% {$GroupType=$_.RecipientTypeDetails;$Name=$_.Name;$SMTP=$_.PrimarySmtpAddress ;Get-DistributionGroupMember -Identity $Name -ResultSize Unlimited | Select-Object @{name=”GroupType”;expression={$GroupType}},@{name=”Group”;expression={$name}},@{name=”GroupSMTP”;expression={$SMTP}},@{name="SMTPDomain";expression={($SMTP).Split("@",2) | Select-Object -Index 1}},@{Label="Member";Expression={$_.Name}},@{Label="MemberSMTP";Expression={$_.PrimarySmtpAddress}},@{Label="MemberType";Expression={$_.RecipientTypeDetails}}} | Export-Csv C:\temp\dg\distributiongroups-and-members.csv –NoTypeInformation
