﻿#Getting Exchange/Office 365 Distribution Group Information

$SaveLocation = Read-Host "Please Enter the Drive Letter or UNC Folder Path where you would like to Save the Outputs"
$groups | Select-Object RecipientTypeDetails,Name,Alias,DisplayName,PrimarySmtpAddress,@{name="SMTP Domain";expression={$_.PrimarySmtpAddress.Domain}},MemberJoinRestriction,MemberDepartRestriction,RequireSenderAuthenticationEnabled,@{Name="ManagedBy";Expression={$_.ManagedBy -join “;”}},@{name=”AcceptMessagesOnlyFrom”;expression={$_.AcceptMessagesOnlyFrom -join “;”}},@{name=”AcceptMessagesOnlyFromDLMembers”;expression={$_.AcceptMessagesOnlyFromDLMembers -join “;”}},@{name=”AcceptMessagesOnlyFromSendersOrMembers”;expression={$_.AcceptMessagesOnlyFromSendersOrMembers -join “;”}},@{name=”ModeratedBy”;expression={$_.ModeratedBy -join “;”}},@{name=”BypassModerationFromSendersOrMembers”;expression={$_.BypassModerationFromSendersOrMembers -join “;”}},@{Name="GrantSendOnBehalfTo";Expression={$_.GrantSendOnBehalfTo -join “;”}},ModerationEnabled,SendModerationNotifications,@{Name="EmailAddresses";Expression={$_.EmailAddresses -join “;”}} | Export-Csv "$SaveLocation\distributiongroups.csv" -NoTypeInformation
$groupSMTP | Get-DistributionGroup -ResultSize Unlimited | Select-Object RecipientTypeDetails,PrimarySmtpAddress -ExpandProperty emailaddresses | select RecipientTypeDetails,PrimarySmtpAddress, @{name="SMTPALIAS";expression={$_}} | Export-Csv "$SaveLocation\distributiongroups-SMTPproxy.csv" -NoTypeInformation
$GroupMembers | $groups |% {$GroupType=$_.RecipientTypeDetails;$Name=$_.Name;$SMTP=$_.PrimarySmtpAddress ;Get-DistributionGroupMember -Identity $Name | Select-Object @{name=”GroupType”;expression={$GroupType}},@{name=”Group”;expression={$name}},@{name=”GroupSMTP”;expression={$SMTP}},@{name="SMTPDomain";expression={($SMTP).Split("@",2) | Select-Object -Index 1}},@{Label="Member";Expression={$_.Name}},@{Label="MemberSMTP";Expression={$_.PrimarySmtpAddress}},@{Label="MemberType";Expression={$_.RecipientTypeDetails}}} | Export-Csv "$SaveLocation\distributiongroups-and-members.csv" –NoTypeInformation
ii $SaveLocation
Exit