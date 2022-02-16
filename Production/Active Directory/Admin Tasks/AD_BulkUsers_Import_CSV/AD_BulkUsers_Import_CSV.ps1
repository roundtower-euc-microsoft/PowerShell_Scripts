#Louisiana Board of Regents User Import Script
#By Corey St. Pierre, Geocent LLC
#Change Date 10/15/2015
#
#Step 1 - Ask if there are any users to import
#In this step, the script will prompt a user dialog box with a yes and no button, asking if there are users to import
#If Yes is chosen, the script proceeds. If No is chosen, the script ends
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$caption = "Warning!"
$message = "Do you have any users to import today? Yes or No?"
$yesNoButtons = 4

if ([System.Windows.Forms.MessageBox]::Show($message, $caption, $yesNoButtons) -eq "NO") {
"You answered no"
exit
}
else {
"You answered yes"
}
#Step 2 - Import Proper PowerShell Modules
#This step will import the Active Directory and Quest AD Management PS modules
Import-Module ActiveDirectory
add-PSSnapin  quest.activeroles.admanagement
#Step 3 - Create a File Selection Box to import the CSV
#In this step, an Open File dialog box will be created, so that users can choose the location of their CSV, and/or copy and paste it into an RDP session
Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}
#Step 4 - Import CSV Information
#In this step, the information in the CSV that was imported will be translated into the Active Directory attributes for the users to be created
#The information will be parsed into variables that will be used to set the appropriate attributes. The users will be created first, and then added to a group specified in the CSV.
$CSVFile = Get-FileName C:\Scripts\CSVFile
$Users = Import-Csv $CSVFile         
foreach ($User in $Users)            
{            
    $Displayname = $User.'Firstname' + " " + $User.'Lastname'            
    $UserFirstname = $User.'Firstname'            
    $UserLastname = $User.'Lastname'            
    $OU = $User.'OU'            
    $SAM = $User.'SAM'            
    $UPN = $User.'SAM' + "@" + $User.'Maildomain'            
    $Description = $User.'Description'
    $Title = $User.'Title'
    $Email = $User.'EmailAddress'
    $Telephone = $User.'OfficePhone'
    $Organization = $User.'Organization'
    $Password = $User.'Password'
    $Group = $User.'MemberOf'
    New-ADUser -Name "$Displayname" -DisplayName "$Displayname" -SamAccountName $SAM -UserPrincipalName $UPN -GivenName "$UserFirstname" -Surname "$UserLastname" -Description "$Description"  -Title "$Title" -EmailAddress "$Email" -OfficePhone "$Telephone" -Organization "$Organization" -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Enabled $true -Path "$OU" 
    Add-QADGroupMember $Group -member $SAM
}





