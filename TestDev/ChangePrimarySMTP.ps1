# Run under Exchange Powershell
#
#
$errlist = @()
#
if($False -eq (Test-path -path ".\MigrationConfig.json" -PathType leaf)) {
    write-host "Configuration file not found"
    write-host "Ensure MigrationConfig.json is in the current folder"
    exit
}
$Config = get-content ".\MigrationConfig.json" | convertfrom-json
$CustomerShortName = $Config.CustomerShortName
$ErrorLogBaseName = "ErrorLog" + $CustomerShortName
$LogName = "Log" + $CustomerShortName + ".txt"
$OldPrimarySMTP = $Config.OldPrimarySMTP
$NewPrimarySMTP = $Config.NewPrimarySMTP
$SMTPRegex = "*@" + $OldPrimarySMTP
$MailBoxes = get-RemoteMailbox -Resultsize unlimited | Where-Object {$_.PrimarySMTPAddress -like $SMTPRegex}
Write-host "Found " $MailBoxes.count
Add-Content -Path $LogName -Value ("Mailboxes found " + $Mailboxes.count)
#
$MailBoxes | Select-Object Name,primarySMTPAddress,alias | export-csv -notypeinformation -path "ChangePrimarySMTPList.csv"
#
#
foreach ($mailbox in $MailBoxes) {
    $primary = $mailbox.PrimarySMTPAddress
    $temp= $primary.split("@")
    $base = $temp[0]
    $new = $base + $NewPrimarySMTP
    try {
        Set-RemoteMailbox $mailbox.Alias -EmailAddressPolicyEnabled $FALSE 
        Set-RemoteMailbox $mailbox.Alias -PrimarySMTPAddress $New 
    }
    catch {
        $errlist += $mailbox
    }
}
$errlist | export-csv -notypeinformation -path ($ErrorLogBaseName + ".csv")