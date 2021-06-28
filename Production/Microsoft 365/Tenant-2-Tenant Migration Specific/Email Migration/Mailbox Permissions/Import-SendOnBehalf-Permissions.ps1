$Permissions = Import-Csv C:\temp\sendonbehalfpermissions.csv

forEach ($Perm in $Permissions)
    {
        $Delegator = $Perm.Delegator
        $Delegate = $Perm.Delegate

        Set-Mailbox $Delegator -GrantSendOnBehalfTo $Delegate
    }