#Script Starting
Write-Host "Script is Starting" -BackgroundColor Black -ForegroundColor Yellow

#Setting your Cloud UPN for Exchange Online
Write-Host "Prompting for your UPN for Exchange Online" -BackgroundColor Black -ForegroundColor Yellow
Start-Sleep 2
$EXOUserCredential = Get-Credential
Write-Host "Exchange Online Credential Set Credential Set. Continuing..." -BackgroundColor Black -ForegroundColor Green

#Adding Visual Basic Assembly for Box Prompts
Add-Type -AssemblyName Microsoft.VisualBasic

		#Checking for Microsoft Exchange Online PowerShell Module
		Write-Host "Checking for Presence of EXO PowerShell Module" -BackgroundColor Black -ForegroundColor Yellow


		$EXOModulePath = ((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter CreateExoPSSession.ps1 -Recurse).FullName | ?{ $_ -notmatch "_none_" } | select -First 1)
			if ($EXOModulePath -ne $null){
				Write-Host "The Exchange Online PowerShell Module is present, continuing script..." -BackgroundColor Black -ForegroundColor Green
				}

			else{
				Write-Host "The Microsoft Exchange Online PowerShell Module is not installed. Redirecting to download page...."
				Start-Process http://aka.ms/exopspreview
				Write-Host "Script is ending. Please rerun after installing the Microsoft Exchange Online PowerShell Module."
				Start-Sleep 3
				Exit
				}

		#Now moving on to the InPlace HOld in Exchange Online. Importing the Module for EXOP.
		Write-Host "Importing the Exchange Online PS Module and Connecting" -BackgroundColor Black -ForegroundColor Yellow
		Import-Module $EXOModulePath
		Connect-EXOPSSession -Credential $EXOUserCredential
		
		#Function to Import CSV file
		Function Get-FileName($initialDirectory)
		{
			[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
			
			$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
			$OpenFileDialog.initialDirectory = $initialDirectory
			$OpenFileDialog.filter = "CSV (*.csv)| *.csv"
			$OpenFileDialog.ShowDialog() | Out-Null
			$OpenFileDialog.filename
		}	
		
		$CSVFile = Get-FileName C:\Scripts\CSVFile
		$ReMboxes = Import-Csv $CSVFile         
		foreach ($ReMbox in $RemBoxes){

		#Setting IPH Variable
		$UPN = $ReMbox.'UserPrincipalName'
		Write-Host "Setting IPH Variable..." -BackgroundColor Black -ForegroundColor Yellow
		$IPHMBNAME = $UPN
		$IPHMBNAME = $IPHMBNAME.Split("@")[0]
		$PolicyName = "$IPHMBNAME" + "_20yrIPH"
		Write-Host "IPH Variables Set, Continuting..." -BackgroundColor Black -ForegroundColor Green
		#Setting the InPlace Hold
		Try
		{
			#Add user to their own In Place Hold
			Write-Host "Enabling $ReMbox for IPH on Policy Name $PolicyName" -BackgroundColor Black -ForegroundColor Yellow
			New-MailboxSearch $PolicyName -SourceMailboxes $UPN -InPlaceHoldEnabled $true -ItemHoldPeriod 7300
			Write-host "$ReMbox has been placed in the $PolicyName hold. Script is ending..." -BackgroundColor Black -ForegroundColor Green
		}

		Catch
		{
			Write-Error "$ReMbox - Could Not set In-pace hold - Error $($Error[0])"
		}
	}
Exit
