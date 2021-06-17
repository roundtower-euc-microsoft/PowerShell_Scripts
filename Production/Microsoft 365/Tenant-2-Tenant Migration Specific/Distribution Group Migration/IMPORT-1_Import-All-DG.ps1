Import-Csv C:\temp\dg\distributiongroups.csv | ForEach-Object{
    $RecipientTypeDetails=$_.RecipientTypeDetails
    $Name=$_.Name
    $Alias=$_.Alias
    $DisplayName=$_.DisplayName
    $smtp=$_.PrimarySmtpAddress
    $RequireSenderAuthenticationEnabled=[System.Convert]::ToBoolean($_.RequireSenderAuthenticationEnabled)
    $join=$_.MemberJoinRestriction
    $depart=$_.MemberDepartRestriction
    $ManagedBy=$_.ManagedBy -split ';'
    $ModeratedBy=$_.ModeratedBy -split ';'
    $AcceptMessagesOnlyFrom=$_.AcceptMessagesOnlyFrom -split ';'
    $AcceptMessagesOnlyFromDLMembers=$_.AcceptMessagesOnlyFromDLMembers -split ';'
    $AcceptMessagesOnlyFromSendersOrMembers=$_.AcceptMessagesOnlyFromSendersOrMembers -split ';'
    
    if ($RecipientTypeDetails -eq "MailUniversalSecurityGroup")
        {
        if ($ManagedBy)
            {
            New-DistributionGroup -Type security -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -MemberJoinRestriction $join -MemberDepartRestriction $depart -ManagedBy $ManagedBy -ModeratedBy $ModeratedBy -OrganizationalUnit "OU=Agro,OU=Distribution Groups,OU=Groups,DC=AMERICOLD,DC=COM"
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $alias -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled
            }
            Else
            {
            New-DistributionGroup -Type security -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -MemberJoinRestriction $join -MemberDepartRestriction $depart -ModeratedBy -OrganizationalUnit "OU=Agro,OU=Distribution Groups,OU=Groups,DC=AMERICOLD,DC=COM"
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $alias -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled
            }
        }

    if ($RecipientTypeDetails -eq "MailUniversalDistributionGroup")
        {
        if ($ManagedBy)
            {
            New-DistributionGroup -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -MemberJoinRestriction $join -MemberDepartRestriction $depart -ManagedBy $ManagedBy -ModeratedBy $ModeratedBy -OrganizationalUnit "OU=Agro,OU=Distribution Groups,OU=Groups,DC=AMERICOLD,DC=COM"
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $alias -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled
            }
            Else
            {
            New-DistributionGroup -Name $Name -Alias $Alias -DisplayName $DisplayName -PrimarySmtpAddress $smtp -MemberJoinRestriction $join -MemberDepartRestriction $depart -OrganizationalUnit "OU=Agro,OU=Distribution Groups,OU=Groups,DC=AMERICOLD,DC=COM"
            Start-Sleep -s 10
            Set-DistributionGroup -Identity $alias -RequireSenderAuthenticationEnabled $RequireSenderAuthenticationEnabled
            }
        }

    if ($AcceptMessagesOnlyFrom) {Set-DistributionGroup -Identity $Name -AcceptMessagesOnlyFrom $AcceptMessagesOnlyFrom}
    if ($AcceptMessagesOnlyFromDLMembers) {Set-DistributionGroup -Identity $Name -AcceptMessagesOnlyFromDLMembers $AcceptMessagesOnlyFromDLMembers}
    if ($AcceptMessagesOnlyFromSendersOrMembers) {Set-DistributionGroup -Identity $Name -AcceptMessagesOnlyFromSendersOrMembers $AcceptMessagesOnlyFromSendersOrMembers}
  }
