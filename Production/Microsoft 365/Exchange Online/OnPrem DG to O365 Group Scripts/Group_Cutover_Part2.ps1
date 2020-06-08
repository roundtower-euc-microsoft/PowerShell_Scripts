#Step 1 - Import Proper PowerShell Modules
#This step will import the Correct Modules and make sure that the ExecutionPolicy is Set Correctly
#
Import-Module ActiveDirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Set-ExecutionPolicy Bypass -Force
#
#Step 2 - Import CSV Information
#In this step, the information in the CSV that was imported will be parsed for the Group Script to be Run
#
$Groups = Import-CSV .\CSVFiles\GroupCSV.csv      
foreach ($Group in $Groups)            
{            
    $GroupName = $Group.'GroupName'
    $GroupContactAlias = $Group.'GroupContactAlias'
    $GroupContactExternalEmail = $Group.'GroupContactExternalEmail'
    $GroupContactOU = $Group.'GroupContactOU'

    Remove-DistributionGroup -Identity $GroupName -Confirm:$false

    Start-Sleep 5

    New-MailContact -ExternalEmailAddress $GroupContactExternalEmail -Name $GroupName -Alias $GroupContactAlias -OrganizationalUnit $GroupContactOU
}