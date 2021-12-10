#
$out = @()
$Groups = get-AzureAdGroup -all $TRUE
foreach($Group in $Groups) {
    $Members = get-AzureADGroupMember -All $TRUE -Objectid $Group.Objectid
    foreach($Member in $Members) {
        $obj = New-Object Object
        Add-Member -InputObject $obj -MemberType NoteProperty -Name GroupDisplayName -Value $Group.DisplayName
        Add-Member -InputObject $obj -MemberType NoteProperty -Name GroupMailNickName -Value $Group.MailNickName
        Add-Member -InputObject $obj -MemberType NoteProperty -Name MemberDisplayName -Value $Member.DisplayName
        Add-Member -InputObject $obj -MemberType NoteProperty -Name MemberObjectType -Value $Member.ObjectType
        Add-Member -InputObject $obj -MemberType NoteProperty -Name MemberUserPrincipalName -Value $Member.UserPrincipalName

        $out += $obj
    }
}
$out | export-csv -notypeinformation -path "CASAllGroupMembers.csv"