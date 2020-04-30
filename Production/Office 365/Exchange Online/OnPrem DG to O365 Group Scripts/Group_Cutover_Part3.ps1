#Step 1 - Import Proper PowerShell Modules
#This step will import the Correct Modules and make sure that the ExecutionPolicy is Set Correctly
#
Import-Module ActiveDirectory
Import-Module MSOnline
Set-ExecutionPolicy Bypass -Force
#
#Step 2 - Connect to Exchange Online
#This step will connect you to Exchange Online. Use your Global Admin Account to connect
#
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session
#
#Step 3 - Import CSV Information
#In this step, the information in the CSV that was imported will be parsed for the Group Script to be Run
#
$Groups = Import-CSV .\CSVFiles\GroupCSV.csv      
foreach ($Group in $Groups)            
{            
    $GroupName = $Group.'GroupName'

    &.\Recreate-DistributionGroup.ps1 -Group $GroupName -Finalize
}