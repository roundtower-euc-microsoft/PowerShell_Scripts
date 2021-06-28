$Permissions = Import-Csv C:\temp\sendaspermissions.csv

forEach ($Perm in $Permissions)
    {
        $Delegator = $Perm.Delegator
        $Delegate = $Perm.Delegate
        $AccessRights = $Perm.AccessRights
                
        Add-RecipientPermission -Identity $Delegator -AccessRights $AccessRights -Trustee $Delegate -Confirm:$false
    }