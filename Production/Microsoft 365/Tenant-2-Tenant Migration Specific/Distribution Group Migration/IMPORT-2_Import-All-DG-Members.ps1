Import-Csv C:\temp\dg\distributiongroups-and-members.csv | ForEach-Object{
$RecipientTypeDetails=$_.GroupType
$GroupSMTP=$_.GroupSMTP
$MemberSMTP=$_.MemberSMTP

    if ($RecipientTypeDetails -eq "MailUniversalSecurityGroup")
        {
        Add-DistributionGroupMember -Identity $GroupSMTP -Member $MemberSMTP -BypassSecurityGroupManagerCheck
        }
    
    if ($RecipientTypeDetails -eq "MailUniversalDistributionGroup")
        {
        Add-DistributionGroupMember -Identity $GroupSMTP -Member $MemberSMTP
        }
}
