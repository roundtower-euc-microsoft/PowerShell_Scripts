#
# Al Lipscomb - 2021
# Basic script to export all SMTP proxy addresses from Exchange Online.
# This script should be run from a PowerShell session already logged in.
# Older version of Exchange present the EmailAddresses objects in a different format. 
#
#
$objs = @()
$r = get-recipient -resultsize unlimited
foreach($obj in $r) {
   $addr = $obj.EmailAddresses
   foreach($a in $addr) {
      if($a -notlike "smtp:*"){
         Continue
         }
      $rec = @{}
      $rec.type = $obj.RecipientType
      $rec.alias = $obj.Alias
      $rec.primary = $obj.PrimarySMTPAddress
      $rec.ou = $obj.OrganizationalUnit
      $rec.name = $obj.name
      $rec.address = $a
      $o = New-Object -TypeName PSObject -Prop $rec
      $objs += $o
      }
}

$objs | export-csv -path ".\SMTPFullReport.csv" -notypeinformation
      
