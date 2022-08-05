#Get all Mailboxes into temp variable
$Mailboxes = Get-Mailbox -ResultSize Unlimited

#Delegates Reports
forEach ($mailbox in $mailboxes)
    {
        Get-Mailbox $mailbox.UserPrincipalName | Get-MailboxPermission | Select Identity, User, Deny, AccessRights, IsInherited| Where {($_.user -ne "NT AUTHORITY\SELF")}| Export-Csv -Path "c:\temp\NonOwnerPermissions.csv" -NoTypeInformation -Append
        Get-Mailbox $mailbox.UserPrincipalName | Get-RecipientPermission| where {($_.trustee -ne "NT AUTHORITY\SELF")}|select Identity,Trustee,AccessControlType,AccessRights,IsInherited  | Export-Csv -Path "c:\temp\sendaspermissions.csv" –NoTypeInformation -Append
        $GrantSendOn= Get-Mailbox  $mailbox.UserPrincipalName | where {($_.GrantSendOnBehalfTo -ne "")} 

        $Out=foreach ($user in $GrantSendOn.GrantSendOnBehalfTo) {

            $obj= New-Object System.Object

            $obj|Add-Member NoteProperty eMail $GrantSendOn.WindowsEmailAddress

            $obj|Add-Member NoteProperty DisplayName $GrantSendOn.DisplayName

            $obj|Add-Member NoteProperty User $user

            $obj }

    $Out| Export-Csv -Path "c:\temp\sendonbehalfpermissions.csv" –NoTypeInformation -Append
}