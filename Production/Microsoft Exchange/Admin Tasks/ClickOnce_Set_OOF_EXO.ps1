##############################################################
##                                                          ##
#        Click-Once Set Out of Office Status Script          #
#                                                            #
#                  By: Corey St. Pierre                      #
#                      Sr. Microsoft Systems Engineer        #
#                      RoundTower Technologies, LLC          #
#                      corey.stpierre@roundtower.com         #
##                                                          ##
##############################################################

<#
    ./SYNOPSIS

        The purpose of this script is to provide a quick,
        efficient way for help desk and IT Admins to grant
        set the Out of Office setting for users that
        request it be set. The script will check for the 
        presence of the Exchange Online PowerShell module 
        (which supports MFA) and will prompt for download and 
        installation in order to  continue if it is not 
        present. After successful verification, prompts for 
        the requesting user's mailbox alias, the message to
        be set, the time frame, and/or if an existing
        setting for OOF needs to be turned off. The script
        will provide a confirmation of the OOF message before
        executing the command.

    ./WARNING
        This script is intended to be used as is. It has 
        been thoroughly tested for its intended purposes.
        Any modifications to this script without the 
        creators consent can result in loss of script
        functionality and/or data loss. The creator is 
        not responisble for any data loss due to misuse
        of said script.

    ./SOURCES
    Deva [MSFT] - How to connect to Exchange Online PowerShell using multi-factor authentication??
    https://blogs.msdn.microsoft.com/deva/2019/03/05/how-to-connect-to-exchange-online-powershell-using-multi-factor-authentication/

    https://gallery.technet.microsoft.com/office/Office-365-Connection-47e03052

    Microsoft TechNet
    https://docs.microsoft.com/en-us/powershell/module/exchange/mailboxes/add-mailboxpermission?view=exchange-ps

#>

#Script Starting
Write-Host "Script is Starting" -BackgroundColor Black -ForegroundColor Yellow
Write-Host "Prompting for your UPN" -BackgroundColor Black -ForegroundColor Yellow

#Adding Visual Basic Assembly for Box Prompts
Add-Type -AssemblyName Microsoft.VisualBasic
$UPN = [Microsoft.VisualBasic.Interaction]::InputBox("Enter your UPN (i.e. user@octapharma.com","$env:UPN")

        #Adding Function - Show message box popup and return the button clicked by the user.
        function Read-MultiLineInputBoxDialog([string]$Message, [string]$WindowTitle, [string]$DefaultText)
        {
                Add-Type -AssemblyName System.Drawing
                Add-Type -AssemblyName System.Windows.Forms
     
                # Create the Label.
                $label = New-Object System.Windows.Forms.Label
                $label.Location = New-Object System.Drawing.Size(10,10) 
                $label.Size = New-Object System.Drawing.Size(280,20)
                $label.AutoSize = $true
                $label.Text = $Message
     
                # Create the TextBox used to capture the user's text.
                $textBox = New-Object System.Windows.Forms.TextBox 
                $textBox.Location = New-Object System.Drawing.Size(10,40) 
                $textBox.Size = New-Object System.Drawing.Size(575,200)
                $textBox.AcceptsReturn = $true
                $textBox.AcceptsTab = $false
                $textBox.Multiline = $true
                $textBox.ScrollBars = 'Both'
                $textBox.Text = $DefaultText
     
                # Create the OK button.
                $okButton = New-Object System.Windows.Forms.Button
                $okButton.Location = New-Object System.Drawing.Size(415,250)
                $okButton.Size = New-Object System.Drawing.Size(75,25)
                $okButton.Text = "OK"
                $okButton.Add_Click({ $form.Tag = $textBox.Text; $form.Close() })
     
                # Create the Cancel button.
                $cancelButton = New-Object System.Windows.Forms.Button
                $cancelButton.Location = New-Object System.Drawing.Size(510,250)
                $cancelButton.Size = New-Object System.Drawing.Size(75,25)
                $cancelButton.Text = "Cancel"
                $cancelButton.Add_Click({ $form.Tag = $null; $form.Close() })
     
                # Create the form.
                $form = New-Object System.Windows.Forms.Form 
                $form.Text = $WindowTitle
                $form.Size = New-Object System.Drawing.Size(610,320)
                $form.FormBorderStyle = 'FixedSingle'
                $form.StartPosition = "CenterScreen"
                $form.AutoSizeMode = 'GrowAndShrink'
                $form.Topmost = $True
                $form.AcceptButton = $okButton
                $form.CancelButton = $cancelButton
                $form.ShowInTaskbar = $true
     
                # Add all of the controls to the form.
                $form.Controls.Add($label)
                $form.Controls.Add($textBox)
                $form.Controls.Add($okButton)
                $form.Controls.Add($cancelButton)
     
                # Initialize and show the form.
                $form.Add_Shown({$form.Activate()})
                $form.ShowDialog() > $null   # Trash the text of the button that was clicked.
     
                # Return the text that the user entered.
                return $form.Tag
            }

#Checking for Microsoft Exchange Online PowerShell Module
Write-Host "Checking for Presence of EXO PowerShell Module" -BackgroundColor Black -ForegroundColor Yellow

$EXOModulePath = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)
    if ($EXOModulePath -ne $null){
        Import-Module $EXOModulePath
        Connect-EXOPSSession -UserPrincipalName $UPN
        }

    else{
        Write-Host "The Microsoft Exchange Online PowerShell Module is not installed. Redirecting to download page...."
        Start-Process http://aka.ms/exopspreview
        Write-Host "Script is ending. Please rerun after installing the Microsoft Exchange Online PowerShell Module."
        Start-Sleep 3
        Exit
        }

#Now Asking for Mailbox that you wish to Delegate Full Access to
#Adding Visual Basic Assembly for Box Prompts
Add-Type -AssemblyName Microsoft.VisualBasic

####BEGIN MULTI USER LOOP#####
$title = 'Set OOF For User'
$msg   = 'Do you want to proceed with enabling/disabling OOf for a user?'

$yes = New-Object Management.Automation.Host.ChoiceDescription '&Yes'
$no  = New-Object Management.Automation.Host.ChoiceDescription '&No'
$options = [Management.Automation.Host.ChoiceDescription[]]($yes, $no)
$default = 1  # $no

do{
		$response = $Host.UI.PromptForChoice($title, $msg, $options, $default)
		if ($response -eq 0) {

$OOFUser = [Microsoft.VisualBasic.Interaction]::InputBox("Enter Email Address of the User Requesting that OOF be set","$env:OOFUser")
Write-Host "OOF User is Set." -BackgroundColor Black -ForegroundColor Green


[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

$caption = "Current OOF Setting"
$message = "Do you need to turn off the OOF Setting for the User Specified?"
$yesNoButtons = 4

if ([System.Windows.Forms.MessageBox]::Show($message, $caption, $yesNoButtons) -eq "NO") {
    "You answered no"

        #Set the Internal OOF Message
        $InternalmultiLineText = Read-MultiLineInputBoxDialog -Message "Please Enter the Internal OOF Message Here." -WindowTitle "Out of Office Text" -DefaultText "I am currently out of office...."
        if ($InternalmultiLineText -eq $null) { Write-Host "You clicked Cancel" }
        else { Write-Host "You entered the following text: $InternalmultiLineText" }

        #After message is set, we will export it to a temp DIR and reimport with the <BR> for the HTML line breaks
        $Internalmultilinetext | Out-File $env:TEMP\InternalMessage.txt
        $InternallBody = (Get-content $env:TEMP\InternalMessage.txt) -join '<BR>'

        #Set the External OOF Message
        $ExternalmultiLineText = Read-MultiLineInputBoxDialog -Message "Please Enter the External OOF Message Here." -WindowTitle "Out of Office Text" -DefaultText "I am currently out of office...."
        if ($ExternalmultiLineText -eq $null) { Write-Host "You clicked Cancel" }
        else { Write-Host "You entered the following text: $ExternalmultiLineText" }

        #After message is set, we will export it to a temp DIR and reimport with the <BR> for the HTML line breaks
        $Externalmultilinetext | Out-File $env:TEMP\ExternalMessage.txt
        $ExternalBody = (Get-content $env:TEMP\ExternalMessage.txt) -join '<BR>'

        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        $caption1 = "OOF Time Range"
        $message1 = "Do you need to set a specified time range for Out of Office?"
        $yesNoButtons1 = 4
 
        if ([System.Windows.Forms.MessageBox]::Show($message1, $caption1, $yesNoButtons1) -eq "NO") {
            "You answered no"

             Write-Host "Now setting $OOFUser's Out of Office Setting..." -BackgroundColor Black -ForegroundColor Yellow
             Set-MailboxAutoReplyConfiguration -Identity $OOFUser -AutoReplyState Enabled -InternalMessage $InternallBody -ExternalMessage $ExternalBody -ExternalAudience All
             Write-Host "$OOFUser has had their Out of Office Setting Set. Ending..." -BackgroundColor Black -ForegroundColor Green

             #Removing Temp files used for message creation
             Remove-Item $env:TEMP\InternalMessage.txt
             Remove-Item $env:TEMP\ExternalMessage.txt
             
             Start-Sleep 3
             }

        else {
             "You answered yes"

              Write-Host "Now Asking for Date Range to Set Out of Office Setting for..." -BackgroundColor Black -ForegroundColor Yellow

              $StartDate = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the OOF Start Date (Format is M/DD/YYYY HH:MM:SS","$env:StartDate")
              Write-Host "OOF Start Date Set" -BackgroundColor Black -ForegroundColor Yellow

              $EndDate = [Microsoft.VisualBasic.Interaction]::InputBox("Enter the OOF End Date (Format is M/DD/YYYY HH:MM:SS","$env:EndDate")
              Write-Host "OOF End Date Set" -BackgroundColor Black -ForegroundColor Yellow

              Write-Host "Now setting $OOFUser's Out of Office Setting..." -BackgroundColor Black -ForegroundColor Yellow
              Set-MailboxAutoReplyConfiguration -Identity $OOFUser -AutoReplyState Scheduled -StartTime "$StartDate" -EndTime "$EndDate" -InternalMessage $InternallBody -ExternalMessage $ExternalBody -ExternalAudience All
              Write-Host "$OOFUser has had their Out of Office Setting Set. Ending..." -BackgroundColor Black -ForegroundColor Green

              #Removing Temp files used for message creation
              Remove-Item $env:TEMP\InternalMessage.txt
              Remove-Item $env:TEMP\ExternalMessage.txt

              Start-Sleep 3
              }
        
         }
else {
     "You answered yes"

      Write-Host "Now Disabling Out of Office of $OOFUser" -BackgroundColor Black -ForegroundColor Yellow
      Set-MailboxAutoReplyConfiguration -Identity $OOFUser -AutoReplyState Disabled
      Write-Host "$OOFUser has had their Out of Office Setting Disabled. Ending..." -BackgroundColor Black -ForegroundColor Green

      Start-Sleep 3
      }
   }
} until ($response -eq 1)
