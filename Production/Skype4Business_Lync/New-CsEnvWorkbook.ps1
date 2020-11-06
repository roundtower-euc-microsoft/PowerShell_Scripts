<#
.SYNOPSIS  
	Gather Lync deployment environment and configuration information.
.DESCRIPTION  
	Versions
	1.0 - Initial version created to test information gathering and processing steps.
	1.5 - Re-Write to more efficiently gather data in custom PSObject.
	2.1 - Add Excel Diagram Drawing, and better sorting to Word report with proper sections and sub-sections.
	3.0 - Environment data collection sub-routine has been rewritten to gather additional info and change data storage method.
	3.2 - Added certificate sections.
	4.1 - Updated and cleaned up Text User Interface, fixed duplicate SIP domain listings.
	5.0 - Re-Write to clean up code and separate data gathering and report building functions.
	5.1 - All scripts have been updated to use the en-US culture during runtime, this should resolve most if not all localization issues and is reset when the script completes.
			Excel - Added Excel based report for Voice Configuration parameters
			Excel - Removed reference to THEMEVAL theme colors as this seemed to cause failures for non en-US Excel installs when creating the site backgrounds.
			Word - Corrected some spelling mistakes.
	5.2 - Updates
			Excel - Fixed typo on site name on line 512 that was causing problems.
			Word - Voice sections with more than 5 columns will not be included due to formatting issues, instead there will be a reference to the Excel workbook.
				Clean up some table formatting and empty cells.
	5.3 - Updates
			Excel - Removed automated download of Excel stencils as the path has changed. Update path to use new 2012_Stencil_121412.vss file.
			Word - Updated to add support for Word templates.
	5.4 - Updates
			Collector - Updated to better support Standard Edition servers and Skype for Business.
			Word - Updated to properly parse software version tables for Skype for Business.
	6.0 - Re-Write to clean up code, PSObject, and reduce run time.
.LINK  
	http://www.emptymessage.com
.EMAIL
	ccook@emptymessage.com
.EXAMPLE
	.\New-CsEnvWorkbook.ps1 -EnvDataFile filename.zip
.INPUTS
	None. You cannot pipe objects to this script.
.PARAMETER EnvDataFile
	The file name of the Lync Data File to be used to create the report.
.PARAMETER Visible
	Set the visibility flag for the Word application while the report is being built.
#>
param(
	[Parameter(Mandatory = $false)]
	[string]$EnvDataFile = $null,
	[Parameter(Mandatory = $false)]
	[bool]$Visible = $true
)

$ErrorActionPreference = "Stop"
$OFS = "`r`n"

#region Variables

# Enable logging to file.
$script:EnableLogging = $false
# Set script log file.
$script:LogFileName = ".\Get-CsEnvironmentInfo.log"
# Set the initial path to the directory containing this script.
$script:CurrentPath = Get-Location | Select -ExpandProperty Path

$Colors = @{
	Green = "RGB(16,91,99)"
	Tan = "RGB(255,250,213)"
	Yellow = "RGB(255,211,78)"
	Orange = "RGB(219,158,54)"
	Red = "RGB(189,73,50)"
	Black = "RGB(0,0,0)"
	DarkGray = "RGB(64,64,64)"
	Gray = "RGB(128,128,128)"
	LightGray = "RGB(192,192,192)"
	White = "RGB(255,255,255)"
}

#endregion Variables

#region GUI_Requirements
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName Microsoft.VisualBasic
#endregion GUI_Requirements

function Get-CsRoles{
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Cluster
	)
	
	$Roles = @()
	if($Cluster.InstalledServices -match "CentralMgmt"){$Roles += "Central Management Server"}
	if(($Cluster.InstalledServices -match "UserServices") -and ($Cluster.InstalledServices -match "Registrar")){$Roles += "Front-End"}
	if($Cluster.InstalledServices -match "ConfServices"){$Roles += "A/V Conferencing"}
	if($Cluster.InstalledServices -match "FileStore"){$Roles += "File Share"}	
	if($Cluster.InstalledServices -match "EdgeServer"){$Roles += "Edge"}
	if($Cluster.InstalledServices -match "MediationServer"){$Roles += "Mediation"}
	if($Cluster.InstalledServices -match "ExternalServer"){$Roles += "Trusted Application"}
	if($Cluster.InstalledServices -match "PstnGateway"){$Roles += "Pstn Gateway"}
	if($Cluster.SqlInstances){$Roles += "SQL Server"}
	$RoleList = $Roles -join ", "
	return $Roles
}

function New-GuiOpenDialog {
	[CmdletBinding(SupportsShouldProcess = $True)]
	param(
		[Parameter(ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[bool]$MultiSelect = $false,
		[Parameter(ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[string]$InitialDirectory,
		[Parameter(ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string]$Filter
	)
	
	$dlgOpenFile = New-Object System.Windows.Forms.OpenFileDialog
	$dlgOpenFile.Multiselect = $MultiSelect
	if($InitialDirectory){$dlgOpenFile.InitialDirectory = $InitialDirectory}
	$dlgOpenFile.Filter = $Filter  
	$dlgOpenFile.showHelp = $true
	$dlgOpenFile.ShowDialog() | Out-Null
	return $dlgOpenFile
}

function New-ExcelWorkbook {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string] $DataFileName,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$CsConfig,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[bool] $Visible = $true
	)
	
	# Set the Excel document filename.
	$ExcelDocFileName = $DataFileName.Replace(".xml",".xlsx")
	Update-Log "Creating diagram: $($ExcelDocFileName)"
	
	# Create a new instance of Microsoft Excel to work with.
	Update-Log "Creating new instance of Excel to work with."
	$script:ExcelApplication = New-Object -ComObject "Excel.Application"
	
	# Create a new blank document to work with and make the Excel application visible.
	Update-Log "Creating new Excel document."
	[int]$TabColor = 20
	$ExcelApplication.Visible = $Visible
	$ExcelWorkbooks = $ExcelApplication.Workbooks
	$ExcelWorkbook = $ExcelWorkbooks.Add()
	$ExcelWorksheets = $ExcelWorkbook.WorkSheets
	$ExcelPage = $ExcelWorksheets.Item(1)
	$CurrentSheetNumber = 1
	$CurrentWorkSheet = $ExcelWorksheets.Item($CurrentSheetNumber)
	
	
	$VoicePolicies = $CsConfig.PolicyData.Voice
	$PolicyTypes = $VoicePolicies | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name
	foreach ($PolicyType in $PolicyTypes){
		if ($CurrentSheetNumber -gt $ExcelWorkSheets.Count){
			$CurrentWorkSheet = $ExcelWorkSheets.Add()
		} else {
			$CurrentWorkSheet = $ExcelWorkSheets.Item($CurrentSheetNumber)
		}
		
		Update-Log "Creating $PolicyType policy sheet."

		$CurrentWorkSheet.Name = "$($PolicyType -creplace "([a-z])([A-Z])", '$1 $2')"
		if ($VoicePolicies.$($PolicyType)){
			[System.Collections.ArrayList]$PolicyAttributes = $VoicePolicies.$($PolicyType) | select -Property * -ExcludeProperty Anchor,Identity,Element,XsAnyAttributes,XsAnyElements | Get-Member -MemberType Properties | Select -Expand Name
			
			$CurrentColumn = 1
			$CurrentRow = 2
			
			[string]$CurrentWorkSheet.Cells.Item(1,1).value() = "Identity"
			foreach ($AttributeName in $PolicyAttributes){
				[string]$CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $AttributeName
				$CurrentRow++
			}
		
			# Reset currently selected cell to Row 1, Column 2
			$CurrentRow = 1
			$CurrentColumn = 2
		
			foreach ($Policy in $VoicePolicies.$($PolicyType)) {
				# Set first row as Identity value.
				[string]$CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = $Policy.Identity
				$CurrentRow++
				foreach ($AttributeName in $PolicyAttributes){
					[string]$CurrentWorkSheet.Cells.Item($CurrentRow,$CurrentColumn).value() = "$($Policy.$($AttributeName))"
					$CurrentRow++
				}
				$CurrentRow = 1
				$CurrentColumn++
			}
			$CurrentWorkSheet.UsedRange.Columns.Autofit() | Out-Null
			$objList = $CurrentWorkSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $CurrentWorkSheet.UsedRange, $null,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes,$null)
			$objList.TableStyle = "TableStyleMedium20"
		} else {
			$CurrentWorkSheet.Cells.Item(1,1).value() = "No Settings Found"
		}
		$CurrentSheetNumber++
	}
	Update-Log "Finished creating workbook, saving changes to document."
	$ExcelWorkbook.SaveAs("$ExcelDocFileName")
	Update-Log "Done."
}

function Open-CsDataFile {
	[CmdletBinding(SupportsShouldProcess = $True)]
	param(
		[Parameter(ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$DataFileName = $null
	)

	# Test if filename that was passed as a commandline option exists; if it does not exist, show the GUI file picker dialog for user to select the file.
	try{
		Test-Path -path "$DataFileName"
	}
	catch{
		# Notify user that the specified file could not be found and that the "Open File Dialog" will be presented.
		Update-Log "**Unable to locate specified data file**" ERROR
		Update-Log "Opening file picker for data file selection" WARNING
		
		# Set the file type filter to Zip and XML files.
		$Filter = "CS Env data (*.zip, *.xml)| *.zip; *.xml"
		# Present the Open File dialog box.
		$DataFile = New-GuiOpenDialog -InitialDirectory $script:CurrentPath -Filter $Filter
		$DataFileName = $DataFile.FileName
	}
	
	try{
		Test-Path -Path "$DataFileName" | Out-Null
	}
	catch{
		Update-Log "Cannot open data file." ERROR
		Exit
	}
	
	$DataFile = Get-ChildItem $DataFileName
	
	# If a Zip file was selected, extract the contents before moving forward.
	if ($DataFileName.EndsWith(".zip")){
		$ShellApp = New-Object -COMObject Shell.Application
		$DataFileZip = $ShellApp.NameSpace("$DataFile")
		$DestinationFolder = $ShellApp.NameSpace("$script:CurrentPath")
		Update-Log "$script:CurrentPath"
		Update-Log "Extracting CS Environment data file to $script:CurrentPath"
		$DestinationFolder.CopyHere($DataFileZip.Items()) | Out-Null
	}
	$script:XmlFileName = $DataFileName.Replace(".zip", ".xml")
	
	Update-Log "Importing CS Environment data file."
	$CsConfig = Import-Clixml "$XmlFileName"
	[xml]$script:XMLTopology = $CsConfig.Topology.XML
	return $CsConfig
	
}

function Update-Log {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true, HelpMessage = "No message Provided.")]
		[ValidateNotNullOrEmpty()]
		[string] $Message,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[ValidateSet("Error", "Warning", "Info")]
		[string] $MessageType = "Info"		
	)
	
	try {
		$LogFileMessage = "{0} : {1} : {2}{3}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $MessageType.ToUpper(), ("  " * $Indent), $Message
		switch ($MessageType) { 
			"Info" {Write-Host -BackgroundColor Black -ForegroundColor Gray "$Message"}
			"Warning" {Write-Host -BackgroundColor Black -ForegroundColor Yellow "$Message"}
			"Error" {Write-Host -BackgroundColor Black -ForegroundColor Red "$Message"}
		}
		 if($script:EnableLogging){$LogFileMessage | Out-File -FilePath $LogFileName -Append}
	}
	catch {
		Throw "Creating Log file '$LogFileName'. The error was: '$_'."
	}
}



Update-Log "Starting report creation."
$CsConfig = Open-CsDataFile -DataFileName $EnvDataFile

New-ExcelWorkbook -DataFileName $script:XmlFileName -CsConfig $CsConfig -Visible $Visible









