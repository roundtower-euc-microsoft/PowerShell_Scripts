#
$Config = get-content ".\config.json" | convertfrom-json
$SourceList = Import-CSV -path $Config.SourceUserFile | Where-Object{$_.UserType -eq "Member"}
$TargetDomain = $Config.TargetDomain
$SourceHash = @{}
#
foreach($SourceRecord in $SourceList) {
    $key = $SourceRecord.mailnickname
    Write-host "Key " $key
    $SourceHash.add($key,$SourceRecord)
}
#
$TargetList = Get-AzureADUser -All $true
$TargetHash = @{}
#
foreach($TargetRecord in $TargetList) {
    $key = $TargetRecord.mailnickname
    $TargetHash.add($key,$TargetRecord)
}

#
$missingFromSource = @()
#
foreach($user in $TargetList) {
    $key = $user.mailnickname
    if($SourceHash.contains($key)) {
        continue
    }
    $missingFromSource += $user
}
#
$missingFromTarget = @()
$missingFromSource = @()
foreach($user in $SourceList){
    $key = $user.mailnickname
    if($TargetHash.contains($key)) {
        continue
    }
    $missingFromTarget += $user
}
#
if($missingFromSource.Count -gt 0) {
    write-host "Missing from Source " $missingFromSource.Count
    $missingFromSource | Format-table DisplayName,UserPrincipalName,mailnickname
}
if($missingFromTarget.Count -gt 0) {
    write-host "Missing from Target " $missingFromTarget.Count
    $missingFromTarget | Format-table DisplayName,UserPrincipalName,mailnickname
#
    $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
    $PasswordProfile.Password = "Qjuazx45#rzm"
    Foreach($User in $missingFromTarget) {
        $UserPrincipalName = $User.mailnickname + "@" + $TargetDomain   
        write-host "Adding User UPN: " $UserPrincipalName
        #new-AzureADUser -DisplayName $user.DisplayName -UserPrincipalName $UserPrincipalName -PasswordProfile $PasswordProfile -accountEnabled $True -mailnickname $user.mailnickname
    }
}

write-host "Done"

