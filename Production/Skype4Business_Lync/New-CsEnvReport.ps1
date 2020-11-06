<#
.SYNOPSIS  
	Gather Lync deployment environment and configuration information.
.DESCRIPTION  
	Versions
	1.0 - Initial version created to test information gathering and processing steps.
	1.5 - Re-Write to more efficiently gather data in custom PSObject.
	2.1 - Add Visio Diagram Drawing, and better sorting to Word report with proper sections and sub-sections.
	3.0 - Environment data collection sub-routine has been rewritten to gather additional info and change data storage method.
	3.2 - Added certificate sections.
	4.1 - Updated and cleaned up Text User Interface, fixed duplicate SIP domain listings.
	5.0 - Re-Write to clean up code and separate data gathering and report building functions.
	5.1 - All scripts have been updated to use the en-US culture during runtime, this should resolve most if not all localization issues and is reset when the script completes.
			Excel - Added Excel based report for Voice Configuration parameters
			Visio - Removed reference to THEMEVAL theme colors as this seemed to cause failures for non en-US Visio installs when creating the site backgrounds.
			Word - Corrected some spelling mistakes.
	5.2 - Updates
			Visio - Fixed typo on site name on line 512 that was causing problems.
			Word - Voice sections with more than 5 columns will not be included due to formatting issues, instead there will be a reference to the Excel workbook.
				Clean up some table formatting and empty cells.
	5.3 - Updates
			Visio - Removed automated download of Visio stencils as the path has changed. Update path to use new 2012_Stencil_121412.vss file.
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
	.\New-CsEnvReport.ps1 -EnvDataFile filename.zip
.INPUTS
	None. You cannot pipe objects to this script.
.PARAMETER EnvDataFile
	The file name of the Lync Data File to be used to create the report.
.PARAMETER Visible
	Set the visibility flag for the Word application while the report is being built.
.PARAMETER Template
	The file name of the Word document template to be used to create the report.
#>
param(
	[Parameter(Mandatory = $false)]
	[string]$EnvDataFile = $null,
	[Parameter(Mandatory = $false)]
	[bool]$Visible = $true,
	[Parameter(Mandatory = $false)]
	[string]$Template = $null
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


# Word document template style setting names.
$script:TableStyleName = "Grid Table 3 - Accent 1"
$script:TitleStyleName = "Title"
$script:NormalStyleName = "Normal"
$script:Heading1StyleName = "Heading 1"
$script:Heading2StyleName = "Heading 2"
$script:Heading3StyleName = "Heading 3"
$script:Heading4StyleName = "Intense Emphasis"

#endregion Variables

#region GUI_Requirements
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type -AssemblyName Microsoft.VisualBasic
#endregion GUI_Requirements

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
	return $CsConfig
	
}

function New-WordReport {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string] $DataFileName,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$CsConfig,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Template,
		[Parameter(Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[bool] $Visible
	)
	
	# Set the Word document filename.
	$WordDocFileName = $DataFileName.Replace(".xml", ".docx")
	Update-Log "Creating report: $($WordDocFileName)"
	
	# Create new instance of Microsoft Word to work with.
	Update-Log "Creating new instance of Word to work with."
	$script:WordApplication = New-Object -ComObject "Word.Application"
	
	# Create a new blank document to work with or open template if one is specified and make the Word application visible.
	if ($Template){
		$script:WordDocument = $script:WordApplication.Documents.Open("$($Template.FullName)")
	} else {
		$script:WordDocument = $script:WordApplication.Documents.Add()
	}
	$script:WordApplication.Visible = $Visible
	$script:WordDocument.SaveAs([ref]$WordDocFileName)
	
	# Word refers to the current cursor location as it's selection.
	$script:Selection = $script:WordApplication.Selection	
	# Formatting and document navigation commands. These functions must be defined this way as the selection will not exist until runtime.
	# The New-Line function is the equivalent of pushing the Enter key for a new line.
	New-Item function:script:New-Line -Value {param([int]$Count = 1);for ($i = 0; $i -lt $Count; $i++){$script:Selection.TypeParagraph()}} -Force| Out-Null
	# The New-PageBreak function inserts a page break.
	New-Item function:script:New-PageBreak -Value {param([int]$Count = 1);for ($i = 0; $i -lt $Count; $i++){$script:Selection.InsertNewPage()}} -Force| Out-Null
	# The MoveTo-End function moves the selection/cursor to the end of the document.
	New-Item function:script:MoveTo-End -Value {$script:Selection.Start = $script:Selection.StoryLength - 1} -Force| Out-Null
	
	# Create cover page for report.
	Update-Log "Creating report cover page."
	New-Line -Count 5
	$Selection.Style = $script:TitleStyleName
	$Selection.ParagraphFormat.Alignment = 1
	$Selection.TypeText("Skype for Business Environment Report")
	New-Line -Count 2
	$Selection.ParagraphFormat.Alignment = 1
	$Selection.Font.Size = 24
	$Selection.TypeText($CsConfig.EnvironmentData.AdDomain.Forest.Name)
	New-Line -Count 2
	$Selection.ParagraphFormat.Alignment = 1
	$Selection.Font.Size = 18
	$Selection.TypeText("Data Gathered: $($CsConfig.TimeStamp)")
	New-PageBreak
	Update-Log "Done."
	MoveTo-End
	
	# Create a blank second page that will hold the Table of Contents. The number specifies how many blank pages we want to insert.
	New-PageBreak
	
	# Create the Topology section of the report.
	Update-Log "Creating Topology report section."
	New-WordHeading -Label "Topology and Architecture" -Style $script:Heading1StyleName
	New-WordHeading -Label "Deployment Summary"-Style $script:Heading2StyleName
	
	# Create a data table for the deployment summary information and then pass it on to the function to create the table in Word.
	$dt = New-Object System.Data.Datatable
	[void]$dt.Columns.Add("Label")
	[void]$dt.Columns.Add("Value")
	$PoolCount = $CsConfig.Topology.Object.Clusters | Where {$_.Machines.Count -gt 1} | measure | select -ExpandProperty Count
	$MachineCount = $CsConfig.Topology.Object.Clusters.Machines.Count
	$SipDomainCount = $CsConfig.EnvironmentData.SipDomains | measure | select -ExpandProperty Count
	[void]$dt.Rows.Add("Total Sites",$CsConfig.Topology.Object.Sites.Count)
	[void]$dt.Rows.Add("Total Pools",$PoolCount)
	[void]$dt.Rows.Add("Total Machines",$MachineCount)
	[void]$dt.Rows.Add("Total User Count",$CsConfig.EnvironmentData.UserData.Count)
	[void]$dt.Rows.Add("Total SIP Domains",$SipDomainCount)
	# Call function to create Word table from DataTable object.
	New-WordTable -DataTable $dt | Out-Null
	MoveTo-End
	New-Line

	# Create heading for sites section.
	New-WordHeading -Label "Sites" -Style $script:Heading2StyleName

	# Enumerate sites and create the report section for each, this includes their pools and machines.
	foreach ($Site in $CsConfig.Topology.Object.Sites){
		New-WordSiteSection $Site
	}
	
	#Create table for internal DNS records.
	Update-Log "Creating Internal DNS record table."
	New-WordHeading -Label "Internal DNS Records" -Style $script:Heading2StyleName
	$dt = New-Object System.Data.Datatable
	[void]$dt.Columns.Add("Label")
	[void]$dt.Columns.Add("Value")
	foreach($Record in $CsConfig.EnvironmentData.InternalDNS.GetEnumerator()){[void]$dt.Rows.Add($Record.Name,$Record.Value)}
	New-WordTable -DataTable $dt | Out-Null
	MoveTo-End
	New-Line

	#Create table for external DNS records.
	Update-Log "Creating External DNS record table."
	New-WordHeading -Label "External DNS Records" -Style $script:Heading2StyleName
	$dt = New-Object System.Data.Datatable
	[void]$dt.Columns.Add("Label")
	[void]$dt.Columns.Add("Value")
	foreach($Record in $CsConfig.EnvironmentData.ExternalDNS.GetEnumerator()){[void]$dt.Rows.Add($Record.Name,$Record.Value)}
	New-WordTable -DataTable $dt | Out-Null
	MoveTo-End
	New-Line

	#Create table for SIP Domains.
	Update-Log "Creating SIP Domain table."
	New-WordHeading -Label "SIP Domains" -Style $script:Heading2StyleName
	$dt = New-Object System.Data.Datatable
	[void]$dt.Columns.Add("Value")
	foreach($SIPDomain in $CsConfig.EnvironmentData.SipDomains){[void]$dt.Rows.Add($SIPDomain.Name)}
	New-WordTable -DataTable $dt | Out-Null
	MoveTo-End
	New-Line

	#Create table for Simple URLs.
	Update-Log "Creating Simple URLs table."
	New-WordHeading -Label "Simple URLs" -Style $script:Heading2StyleName
	$dt = New-Object System.Data.Datatable
	[void]$dt.Columns.Add("Component")
	[void]$dt.Columns.Add("Domain")
	[void]$dt.Columns.Add("URL")
	[void]$dt.Rows.Add("Component","Domain","URL")
	foreach($SimpleURL in $CsConfig.EnvironmentData.SimpleUrls.SimpleUrl){[void]$dt.Rows.Add($SimpleURL.Component,$SimpleURL.Value,$SimpleURL.ActiveUrl)}
	New-WordTable -DataTable $dt -HeaderRow $true| Out-Null
	MoveTo-End
	New-Line

	#Create table for CMS Configuration.
	Update-Log "Creating CMS Configuration table."
	New-WordHeading -Label "CMS Configuration" -Style $script:Heading2StyleName
	$dt = New-Object System.Data.Datatable
	[void]$dt.Columns.Add("Label")
	[void]$dt.Columns.Add("Value")
	$CMSConfig = $CsConfig.EnvironmentData.CMS
	[void]$dt.Rows.Add("Last Updated",$CMSConfig.CMSStatus.LastUpdatedOn)
	[void]$dt.Rows.Add("Active Master FQDN",$CMSConfig.CMSStatus.ActiveMasterFqdn)
	[void]$dt.Rows.Add("Active File Transfer FQDN",$CMSConfig.CMSStatus.ActiveFileTransferAgentFqdn)
	[void]$dt.Rows.Add("StoreProvider",$CMSConfig.ManagementConnection.StoreProvider)
	[void]$dt.Rows.Add("Connection String",$CMSConfig.ManagementConnection.Connection)
	[void]$dt.Rows.Add("Read Only",$CMSConfig.ManagementConnection.ReadOnly)
	[void]$dt.Rows.Add("SQL Server",$CMSConfig.ManagementConnection.SqlServer)
	[void]$dt.Rows.Add("SQL Instance",$CMSConfig.ManagementConnection.SqlInstance)
	[void]$dt.Rows.Add("Mirror SQL Server",$CMSConfig.ManagementConnection.MirrorSqlServer)
	[void]$dt.Rows.Add("Mirror SQL Instance",$CMSConfig.ManagementConnection.MirrorSqlInstance)
	New-WordTable -DataTable $dt| Out-Null
	MoveTo-End
	New-Line

	#Create table for CMS Replication Status.
	Update-Log "Creating CMS Replication Status table."
	New-WordHeading -Label "CMS Replication Status" -Style $script:Heading2StyleName
	$dt = New-Object System.Data.Datatable
	[void]$dt.Columns.Add("Label")
	[void]$dt.Columns.Add("Value")
	$ReplicationStatus = $CsConfig.EnvironmentData.CMS.ReplicationStatus
	foreach($Server in $ReplicationStatus){
		$ReplicationInfo = @()
		$ReplicationInfo += "UpToDate`t`t: $($Server.UpToDate)"
		$ReplicationInfo += "LastStatusReport`t: $($Server.LastStatusReport)"
		$ReplicationInfo += "LastUpdateCreation`t: $($Server.LastUpdateCreation)"
		$ReplicationInfo += "ProductVersion`t: $($Server.ProductVersion)"
		$ServerStatus = $ReplicationInfo -join "`r`n"
		[void]$dt.Rows.Add($Server.ReplicaFqdn,$ServerStatus)
	}
	New-WordTable -DataTable $dt| Out-Null
	MoveTo-End
	New-Line

	if($CsConfig.EnvironmentData.DBMirrorData.Values){
		#Create table for DB Mirror Configuration.
		Update-Log "Creating DB Mirror Configuration table."
		$dt = New-Object System.Data.Datatable
		[void]$dt.Columns.Add("Label")
		[void]$dt.Columns.Add("Value")
		$DBMirrorData = $CsConfig.EnvironmentData.DBMirrorData
		foreach($Pool in $DBMirrorData.GetEnumerator()){
			if($Pool.Value){[void]$dt.Rows.Add($Pool.Name,$Pool.Value)}
		}
		if($dt.Rows){
			New-WordHeading -Label "DB Mirror Configuration" -Style $script:Heading2StyleName
			#New-WordTable -DataTable $dt| Out-Null
		}
		MoveTo-End
		New-Line
	}

	#Create table for Active Directory information.
	Update-Log "Creating Active Directory Information table."
	New-WordHeading -Label "Active Directory" -Style $script:Heading2StyleName
	$dt = New-Object System.Data.Datatable
	[void]$dt.Columns.Add("Label")
	[void]$dt.Columns.Add("Value")
	$ADConfig = $CsConfig.EnvironmentData.AdDomain
	[void]$dt.Rows.Add("AD Forest",$ADConfig.Forest.Name)
	[void]$dt.Rows.Add("AD Domain Functional Level",$ADConfig.DomainMode.Value)
	[void]$dt.Rows.Add("AD Forest Functional Level",$ADConfig.Forest.ForestMode)
	[void]$dt.Rows.Add("Root Domain",$ADConfig.Forest.RootDomain)
	[void]$dt.Rows.Add("PDC Master",$ADConfig.PdcRoleOwner.Name)
	[void]$dt.Rows.Add("RID Master",$ADConfig.RidRoleOwner.Name)
	[void]$dt.Rows.Add("Infrastructure Master",$ADConfig.InfrastructureRoleOwner.Name)
	[void]$dt.Rows.Add("Schema Master",$ADConfig.Forest.SchemaRoleOwner)
	[void]$dt.Rows.Add("Naming Master",$ADConfig.Forest.NamingRoleOwner)
	New-WordTable -DataTable $dt -HeaderRow $true| Out-Null
	MoveTo-End
	New-Line

	Update-Log "Done creating Topology report section."

	# Create the sections for Policy configurations.
	New-WordPolicySection ExternalConfig
	New-WordPolicySection Voice
	New-WordPolicySection Conferencing
	New-WordPolicySection RGS
	New-WordPolicySection CAC
	New-WordPolicySection LIS
	New-WordPolicySection Policy
	
	# Create Table of Contents for document. The number specifies which page to place the Table of Contents on.
	New-WordTableOfContents -PageNumber 2
	
	Update-Log "Finished creating report, saving changes to document."
	$script:WordDocument.SaveAs([ref]$WordDocFileName)
    #$script:WordApplication.Quit()
	Update-Log "Done."
}

function New-WordTableOfContents {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$PageNumber
	)

	# Go back to the beginning of page two.
	[void]$Selection.GoTo(1, 2, $null, $PageNumber)
	Update-Log "Creating Table of Contents."
	New-WordHeading -Label "Table of Contents" $script:Heading1StyleName
	
	# Create Table of Contents for document.
	# Set Range to beginning of document to insert the Table of Contents.
	$TocRange = $Selection.Range
	$useHeadingStyles = $true 
	$upperHeadingLevel = 1 # <-- Heading1 or Title 
	$lowerHeadingLevel = 2 # <-- Heading2 or Subtitle 
	$useFields = $false 
	$tableID = $null 
	$rightAlignPageNumbers = $true 
	$includePageNumbers = $true 
	# to include any other style set in the document add them here 
	$addedStyles = $null 
	$useHyperlinks = $true 
	$hidePageNumbersInWeb = $true 
	$useOutlineLevels = $true 

	# Insert Table of Contents
	$WordTableOfContents = $WordDocument.TablesOfContents.Add($TocRange, $useHeadingStyles, 
	$upperHeadingLevel, $lowerHeadingLevel, $useFields, $tableID, 
	$rightAlignPageNumbers, $includePageNumbers, $addedStyles, 
	$useHyperlinks, $hidePageNumbersInWeb, $useOutlineLevels) 
	$WordTableOfContents.TabLeader = 0	
}

function New-WordPolicySection {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$SectionName
	)
	
	Update-Log "Creating $SectionName policy section."
	New-WordHeading -Label "$SectionName" $script:Heading1StyleName
	
	$PolicyCmdlets = $CsConfig.PolicyData.($SectionName) | Get-Member -MemberType NoteProperty | Select -ExpandProperty Name
	
	foreach($Policy in $PolicyCmdlets){
		Update-Log "Creating $Policy policy table."
		New-WordHeading -Label "$($Policy -creplace "([a-z])([A-Z])", '$1 $2')" $script:Heading2StyleName
		
		if(($SectionName -match "Voice") -and ($CsConfig.PolicyData.$SectionName.$Policy.Count -gt 5)){
			$script:Selection.TypeText("Please see detailed Voice documentation in Excel workbook.")
			New-Line
		} else {
			$Policies = $CsConfig.PolicyData.$SectionName.$Policy
			if($Policies){
				$PolicyAttributes = $Policies | Select -Property * -ExcludeProperty XsAnyElements,XsAnyAttributes,Element,Anchor,Identity | Get-Member -MemberType NoteProperty | select -ExpandProperty Name | Sort
				#$PolicyAttributes = $Policies | Get-Member -MemberType Property | Select -Property * -ExcludeProperty XsAnyElements,XsAnyAttributes,Element,Anchor,Identity -ExpandProperty Name | sort
				
				$dt = New-Object System.Data.Datatable
				[void]$dt.Columns.Add("Identity")
				foreach($Attribute in $PolicyAttributes){[void]$dt.Columns.Add("$Attribute")}
				
				
				foreach($CurrentPolicy in $Policies){
					$PolicyValues = @()
					$PolicyValues += "$($CurrentPolicy.Identity)"
					foreach($Attribute in $PolicyAttributes){
						if($CurrentPolicy.$Attribute){
							$PolicyValues += "$($CurrentPolicy.$Attribute)"
						} else {
							$PolicyValues += " "
						}
					}
					[void]$dt.Rows.Add($PolicyValues)
				}
				New-WordTable -DataTable $dt -HeaderRow $true -FlipData $true | Out-Null
			} else {
				$script:Selection.TypeText("No Settings Found")
			}
			$Policies = $null
			MoveTo-End
			New-Line
		}
	}
	New-PageBreak
}

function New-WordSiteSection {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Site
	)
	Update-Log "Site: $($Site.Name)"
	New-WordHeading -Label "Site: $($Site.Name)" $script:Heading3StyleName
	New-WordHeading -Label "$($Site.Name) Details" $script:Heading4StyleName
	
	# Create a data table for the site summary information and then pass it on to the function to create the table in Word.
	$dt = New-Object System.Data.Datatable
	[void]$dt.Columns.Add("Label")
	[void]$dt.Columns.Add("Value")
	$PoolCount = $Site.Clusters | Where {$_.Machines.Count -gt 1} | measure | select -ExpandProperty Count
	$MachineCount = $Site.Clusters.Machines.Count
	$FrontEndPoolNames = $Site.Clusters | Where {$_.InstalledServices -match "UserServices"} | select -ExpandProperty Fqdn
	$SiteUserCount = 0
	foreach($Pool in $FrontEndPoolNames){
		$SiteUserCount = $SiteUserCount + $($CsConfig.EnvironmentData.UserData | Where {$_.RegistrarPool.FriendlyName -match $Pool} | measure).count
	}

	[void]$dt.Rows.Add("$($Site.Name)","")
	[void]$dt.Rows.Add("Description",$Site.Description)
	[void]$dt.Rows.Add("Kind",$Site.Kind)
	[void]$dt.Rows.Add("Site ID",$Site.SiteId)
	[void]$dt.Rows.Add("Country Code",$Site.CountryCode)
	[void]$dt.Rows.Add("State / Province",$Site.State)
	[void]$dt.Rows.Add("City",$Site.City)
	[void]$dt.Rows.Add("Total Pools",$PoolCount)
	[void]$dt.Rows.Add("Total Machines",$MachineCount)
	[void]$dt.Rows.Add("Total User Count",$SiteUserCount)
	# Call function to create Word table from DataTable object.
	New-WordTable -DataTable $dt -HeaderRow $true | Out-Null
	MoveTo-End
	New-Line
	
	# If there are any multiserver pools then create the relevant pool sections.
	if ($PoolCount -gt 0) {New-WordPoolSection -Site $Site}
	
	# Create the machine sections for all machines in current site.
	New-WordMachineSection -Site $Site
}

function New-WordPoolSection {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Site
	)
	New-WordHeading -Label "Pools" $script:Heading4StyleName
	foreach ($Pool in $Site.Clusters | Where {$_.Machines.Count -gt 1}){
		Update-Log "Pool: $($Pool.Fqdn)"
		# Create a data table for the pool information and then pass it on to the function to create the table in Word.
		$dt = New-Object System.Data.Datatable
		[void]$dt.Columns.Add("Label")
		[void]$dt.Columns.Add("Value")
		$Roles = Get-CsRoles -Cluster $Pool
		$MachineIds = $Pool.Machines
		$Machines = @()
		foreach($MachineId in $MachineIds){
			$Machines += $CsConfig.Topology.Object.Machines | Where {$_.MachineId -match $MachineId} | Select -ExpandProperty Fqdn
		}
		$MachineList = $Machines -join "`r`n"
		$PoolUserCount = $null
		$PoolUserCount = $($CsConfig.EnvironmentData.UserData | Where {$_.RegistrarPool.FriendlyName -match $Pool.Fqdn} | measure).count
		
		$XmlTopology = [xml]$CsConfig.Topology.XML
		$PoolClusterSplit = $Pool.ClusterId -split ":"
		$FileShares = @()
		$FileShareList = $null
		$FileStoreServices = $XmlTopology.Topology.Services.Service | Where {($_.DependsOn.Dependency.ServiceId.RoleName -match "FileStore") -and ($_.InstalledOn.ClusterId.SiteId -match $PoolClusterSplit[0]) -and($_.InstalledOn.ClusterId.Number -match $PoolClusterSplit[1])}
		$FileStoreServiceId = $FileStoreServices.DependsOn.Dependency.ServiceId | Where {$_.RoleName -match "FileStore"} | select -Unique
		foreach($FileStore in $FileStoreServiceId){$FileShares += $CsConfig.Topology.Object.Services | Where {$_.ServiceId -eq "$($FileStore.SiteId)-FileStore-$($FileStore.Instance)"} | Select -ExpandProperty UncPath}
		$FileShareList = $FileShares -join "`r`n"

		$DBServices = $XmlTopology.Topology.Services.Service | Where {($_.DependsOn.Dependency.ServiceId.RoleName -match "Store") -and ($_.InstalledOn.ClusterId.SiteId -match $PoolClusterSplit[0]) -and($_.InstalledOn.ClusterId.Number -match $PoolClusterSplit[1])}
		$DBServiceId = $DBServices.DependsOn.Dependency.ServiceId | Where {$_.RoleName -match "Store"} | select -Unique
		$DBService = $XmlTopology.Topology.Services.Service | Where {($_.ServiceId.SiteId -eq $DBServiceId.SiteId) -and ($_.ServiceId.RoleName -eq $DBServiceId.RoleName) -and ($_.ServiceId.Instance -eq $DBServiceId.Instance)}
		$DBClusterId = $DBService.InstalledOn.SqlInstanceId.ClusterId
		$DBFqdn = $CsConfig.Topology.Object.Clusters | Where {$_.ClusterId -eq "$($DBClusterId.SiteId):$($DBClusterId.Number)"} | Select -ExpandProperty Fqdn

		$WebServices = $XmlTopology.Topology.Services.Service | Where {($_.DependsOn.Dependency.ServiceId.RoleName -match "WebServices") -and ($_.InstalledOn.ClusterId.SiteId -match $PoolClusterSplit[0]) -and($_.InstalledOn.ClusterId.Number -match $PoolClusterSplit[1])}
		$WebServiceId = $FileStoreServices.DependsOn.Dependency.ServiceId | Where {$_.RoleName -match "WebServices"} | select -Unique
		$WebService = $CsConfig.Topology.Object.Services | Where {$_.ServiceId -eq "$($WebServiceId.SiteId)-WebServices-$($WebServiceId.Instance)"}

		[void]$dt.Rows.Add("$($Pool.Fqdn)","")
		[void]$dt.Rows.Add("Cluster ID",$Pool.ClusterId)
		[void]$dt.Rows.Add("Roles",$Roles)
		[void]$dt.Rows.Add("Machines",$MachineList)
		# The following items are only added to the table if the property is applicable.
		if($FileShareList){[void]$dt.Rows.Add("File Share",$FileShareList)}
		if($DBFqdn){[void]$dt.Rows.Add("Associated SQL Server(s)",$DBFqdn)}
		if($WebService.InternalHost){[void]$dt.Rows.Add("Internal Web FQDN",$WebService.InternalHost)}
		if($WebService.ExternalHost){[void]$dt.Rows.Add("External Web FQDN",$WebService.ExternalHost)}
		if($PoolUserCount){[void]$dt.Rows.Add("Total User Count",$PoolUserCount)}
		# Call function to create Word table from DataTable object.
		New-WordTable -DataTable $dt -HeaderRow $true | Out-Null
		MoveTo-End
		New-Line
		
		
	}
}

function New-WordMachineSection {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Site
	)
	New-WordHeading -Label "Machines"-Style $script:Heading4StyleName
	foreach ($Machine in $CsConfig.Topology.Object.Machines | Where {$_.MachineId -match "$($Site.SiteId):"}){
		Update-Log "Machine: $($Machine.Fqdn)"
		# Grab server data from file.
		$ServerData = $CsConfig.EnvironmentData.ServerData."$($Machine.Fqdn)"
		# Create a data table for the machine information and then pass it on to the function to create the table in Word.
		$dt = New-Object System.Data.Datatable
		[void]$dt.Columns.Add("Label")
		[void]$dt.Columns.Add("Value")
		$Roles = Get-CsRoles -Cluster $Machine.Cluster
		
		$NetInterfaces = $null
		$NetInterfaces = $XmlTopology.Topology.Clusters.Cluster.Machine | Where {$_.Fqdn -match $Machine.Fqdn} | Select -ExpandProperty NetInterface
		
		$CPUCores = $null
		$CPUCores = $ServerData.CPUInfo.Count
		$CPUName = $null
		$CPUName = $ServerData.CPUInfo.Name | Select -Unique
		$CPULoad = $null
		$CPULoad = $ServerData.CPUInfo.LoadPercentage | measure -Average | select -ExpandProperty Average
		
		$TotalRAM = "{0:N0}" -f $($ServerData.OsInfo.TotalVisibleMemorySize / 1024)
		$UsedRAM = "{0:N0}" -f $($($ServerData.OsInfo.TotalVisibleMemorySize / 1024) - $($ServerData.OsInfo.FreePhysicalMemory) / 1024)
		
		$SWVersionList = $null
		$SWVersionTable = @()
		$ServerData.SWVersionInfo | Where {($_.Name -match "Lync") -or ($_.Name -match "Skype")} | foreach {$SWVersionTable += "$($_.Name), $($_.Version)"}
		$SWVersionList = $SWVersionTable -join "`r`n"
		
		[void]$dt.Rows.Add("$($Machine.Fqdn)","")
		[void]$dt.Rows.Add("Machine ID",$Machine.MachineId)
		[void]$dt.Rows.Add("Parent Pool",$Machine.Cluster.Fqdn)
		[void]$dt.Rows.Add("Role(s)",$Roles)
		# The following items are only added to the table if the property is applicable.
		if($NetInterfaces){[void]$dt.Rows.Add("IP Address(es)",$NetInterfaces)}
		if($CPUName){[void]$dt.Rows.Add("CPU Type",$CPUName)}
		if($CPUCores){[void]$dt.Rows.Add("CPU Core(s)",$CPUCores)}
		if($CPULoad){[void]$dt.Rows.Add("CPU Load Percentage","$CPULoad%")}
		if($ServerData.OsInfo.TotalVisibleMemorySize){[void]$dt.Rows.Add("RAM Utilization (Used / Total)","$UsedRAM MB / $TotalRAM MB")}
		if($SWVersionList){[void]$dt.Rows.Add("Software Versions",$SWVersionList)}
		# Call function to create Word table from DataTable object.
		New-WordTable -DataTable $dt -HeaderRow $true | Out-Null
		MoveTo-End
		New-Line
		
		# Create certificate table if needed.
		if(($ServerData.CertificateInfo -ne "ERROR") -and ($ServerData.CertificateInfo)){
			Update-Log "Certificates: $($Machine.Fqdn)"
			# Create a data table for the machine information and then pass it on to the function to create the table in Word.
			$dt = New-Object System.Data.Datatable
			[void]$dt.Columns.Add("Label")
			[void]$dt.Columns.Add("Value")
			
			[void]$dt.Rows.Add("$($Machine.Fqdn) Certificates","")
			foreach($Certificate in $ServerData.CertificateInfo){
				$CertificateTable = @()
				
				$FullSubject = $Certificate.Subject
				$SplitSubject = $FullSubject.Split(",")
				$ShortSubject = $SplitSubject[0].Substring(3, ($SplitSubject[0].Length - 3))

				$CertificateTable += "Subject : $ShortSubject"
				$CertificateTable += "Created : $($Certificate.NotBefore)"
				$CertificateTable += "Expires : $($Certificate.NotAfter)"
				$CertificateTable += "Issuer : $($Certificate.Issuer)"
				$CertificateTable += "Serial Number : $($Certificate.SerialNumber)"
				$CertificateTable += "Thumbprint : $($Certificate.Thumbprint)"
				if($Certificate.AlternativeNames){$CertificateTable += "SAN Names : `r`n$($Certificate.AlternativeNames.Replace(" ", "`r`n"))"}
				$CertificateDetails = $CertificateTable -join "`r`n"
				
				[void]$dt.Rows.Add("$($Certificate.Use)","$CertificateDetails")
			}
			# Call function to create Word table from DataTable object.
			New-WordTable -DataTable $dt -HeaderRow $true | Out-Null
			MoveTo-End
			New-Line
		}
	}
}

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
	return $RoleList
}

function New-WordHeading{
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[string] $Label,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$Style
	)
	$script:Selection.Style = $Style
	$script:Selection.TypeText($Label)
	$script:Selection.TypeParagraph()
	$script:Selection.ClearFormatting()
}

function New-WordTable {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$DataTable,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[bool]$HeaderRow = $false,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[bool]$FlipData = $false
	)
	
	if($FlipData){
		
		# The FlipData parameter flips rows and columns.
		$NumCols = $DataTable.Rows.Count
		$NumRows = $DataTable.Columns.Count

		# If HeaderRow is true then add an additional row to the table for the labels.
		if($HeaderRow){$NumCols = $DataTable.Rows.Count + 1}
		
		$NewTable = $WordDocument.Tables.Add($script:Selection.Range, $NumRows, $NumCols)
		$NewTable.AllowAutofit = $true
		$NewTable.AutoFitBehavior(2)
		$NewTable.AllowPageBreaks = $false
		$NewTable.Style = $script:TableStyleName
		$NewTable.ApplyStyleHeadingRows = $HeaderRow
		
		# Populate data from DataTable object into Word table.
		[string]$NewTable.Cell(1, 1).Range.Text = "Identity"
		for($Row = 2; $Row -le $DataTable.Columns.ColumnName.Count; $Row++){
			[string]$NewTable.Cell($Row, 1).Range.Text = $DataTable.Columns.ColumnName[$Row - 1]
		}

		for($Column = 2; $Column -le $NumCols; $Column++){
			[string]$NewTable.Cell(1, $Column).Range.Text = $DataTable.Rows[$($Column - 2)].Identity
			for($Row = 2; $Row -le $NumRows; $Row++){
				[string]$NewTable.Cell($Row, $Column).Range.Text = $DataTable.Rows[$($Column - 2)].Item($($Row - 1))
			}
		}
	} else {
		$NumCols = $DataTable.Columns.Count
		$NumRows = $DataTable.Rows.Count
		# If HeaderRow is true then add an additional row to the table for the labels.
		if($HeaderRow){$Rows = $DataTable.Rows.Count + 1}
		$NewTable = $WordDocument.Tables.Add($script:Selection.Range, $NumRows, $NumCols)
		$NewTable.AllowAutofit = $true
		$NewTable.AutoFitBehavior(2)
		$NewTable.AllowPageBreaks = $false
		$NewTable.Style = $script:TableStyleName
		$NewTable.ApplyStyleHeadingRows = $HeaderRow
		# Populate data from DataTable object into Word table.
		for($Row = 1; $Row -le $NumRows; $Row++){
			for($Column = 1; $Column -le $NumCols; $Column++){
				[string]$NewTable.Cell($Row, $Column).Range.Text = $DataTable.Rows[$($Row - 1)].Item($($Column - 1))
			}
		}
	}

	return $NewTable
}



Update-Log "Starting report creation."
$CsConfig = Open-CsDataFile -DataFileName $EnvDataFile
if($Template){$DocumentTemplate = Get-Item "$Template"}

New-WordReport -DataFileName $script:XmlFileName -CsConfig $CsConfig -Template $DocumentTemplate -Visible $Visible









