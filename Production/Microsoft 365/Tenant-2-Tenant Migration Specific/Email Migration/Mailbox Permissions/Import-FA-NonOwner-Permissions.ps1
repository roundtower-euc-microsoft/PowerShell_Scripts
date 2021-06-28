$Permissions = Import-Csv C:\temp\NonOwnerPermissions.csv

forEach ($Perm in $Permissions)
    {
        $Delegator = $Perm.Delegator
        $Delegate = $Perm.Delegate
        $AccessRights = $Perm.AccessRights
                
        Remove-MailboxPermission -Identity $Delegator -User $Delegate -AccessRights $AccessRights -Confirm:$False
        Add-MailboxPermission -Identity $Delegator -User $Delegate -AccessRights $AccessRights -InheritanceType All -AutoMapping $true
    }