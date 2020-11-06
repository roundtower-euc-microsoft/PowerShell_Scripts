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
	.\New-CsEnvDiagram.ps1 -EnvDataFile filename.zip
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

$ExternalEdgeFirewallRules = @"
HTTP (TCP) / PORT 80,Outbound,AE
HTTPS (TCP) / PORT 443,Outbound,AE
DNS (TCP-UDP) / PORT 53,Outbound,AE
SIP (TLS) / PORT 443,Inbound,AE
HTTP (SSL) / PORT 4443,Inbound,AE
SIP (MTLS) / PORT 5061,Both,AE
XMPP (TLS) / PORT 5269,Both,AE
PSOM (TLS) / PORT 443,Inbound,WC
STUN (TCP) / PORT 443,Both,AV
STUN (UDP) / PORT 3478,Both,AV
RTP (TCP) / PORT 50000 - 59999,Both,AV
RTP (UDP) / PORT 50000 - 59999,Both,AV
"@
$InternalEdgeFirewallRules = @"
SIP (MTLS) / PORT 5061,Both,FE
PSOM-MTLS (TCP) / PORT 8057,Inbound,FE
SIP-MTLS (TCP) / PORT 5062,Inbound,FE
HTTPS (SSL) / PORT 4443,Inbound
XMPP-MTLS (TCP) / PORT 23456,Inbound,FE
CLS-MTLS (TCP) / 50001 - 500003,Inbound,Any
STUN (TCP) / PORT 443,Inbound,Any
STUN (UDP) / PORT 3478,Inbound,Any
RTP (TCP) / PORT 50000 - 59999,Inbound,Any
RTP (UDP) / PORT 50000 - 59999,Inbound,Any
"@

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

function Add-ConnectorToPage {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$BeginX,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$BeginY,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$EndX,
		[Parameter(Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$EndY,
		[Parameter(Position = 4, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Name = $null,
		[Parameter(Position = 5, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Page = $CurrentPage,
		[Parameter(Position = 6, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Color,
		[Parameter(Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LineWeight = "0.75 pt",
		[Parameter(Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LinePattern = 1,
		[Parameter(Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LineType = 0,
		[Parameter(Position = 10, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$BeginArrow = 0,
		[Parameter(Position = 11, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$EndArrow = 0,
		[Parameter(Position = 12, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$BeginArrowSize = 1,
		[Parameter(Position = 13, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$EndArrowSize = 1,
		[Parameter(Position = 14, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$ConnectorRouteStyle = 16,
		[Parameter(Position = 15, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$ConnectorReroute = 0,
		[Parameter(Position = 16, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$ConnectorShape = 1
	)
	
	$ConnectorOnPage = $Page.Drop($($VisioStencils.Connector), 10, 10)
	if($Name){$ConnectorOnPage.Name = $Name}
	Update-Log "Adding connector $Name to page."

	$ConnectorOnPage.CellsU("BeginX").Formula = "$BeginX"
	$ConnectorOnPage.CellsU("BeginY").Formula = "$BeginY"
	$ConnectorOnPage.CellsU("EndX").Formula = "$EndX"
	$ConnectorOnPage.CellsU("EndY").Formula = "$EndY"
	$ConnectorOnPage.CellsU("LinePattern").Formula = "$LinePattern"
	$ConnectorOnPage.CellsU("CompoundType").Formula = "$LineType"
	$ConnectorOnPage.CellsU("LineWeight").Formula = "$LineWeight"
	$ConnectorOnPage.CellsU("LineColor").Formula = "THEMEGUARD($Color)"
	$ConnectorOnPage.CellsU("BeginArrow").Formula = "$BeginArrow"
	$ConnectorOnPage.CellsU("EndArrow").Formula = "$EndArrow"
	$ConnectorOnPage.CellsU("BeginArrowSize").Formula = "$BeginArrowSize"
	$ConnectorOnPage.CellsU("EndArrowSize").Formula = "$EndArrowSize"
	$ConnectorOnPage.CellsU("ShapeRouteStyle").Formula = "$ConnectorRouteStyle"
	$ConnectorOnPage.CellsU("ConFixedCode").Formula = "$ConnectorReroute"
	$ConnectorOnPage.CellsU("ConLineRouteExt").Formula = "$ConnectorShape"
	$ConnectorOnPage.Text = ""
	
	return $ConnectorOnPage
}

function Add-InternalServerToPage {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$X1,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Y1,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Page = $CurrentPage,
		[Parameter(Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$Shape,
		[Parameter(Position = 4, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Name,
		[Parameter(Position = 5, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Roles,
		[Parameter(Position = 6, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Members = $null
	)
	
	$ShapeOnPage = $Page.Drop($($VisioStencils."$Shape"), 10, 10)
	
	$ShapeOnPage.Name = $Name
	$ShapeOnPage.CellsU("PinX").Formula = "$X1"
	$ShapeOnPage.CellsU("PinY").Formula = "$Y1"
	$ShapeOnPage.Text = ""
	Update-Log "Adding server $Name to page."
	
	$FqdnLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 ($X1 + 2.5) -Y1 ($Y1 + 0.5) -Name "$($Name)FqdnLabel" -Height "0.25 in" -Width 4 -Color $($Colors.White) -Transparency 100 -LineWeight "0 pt"
	Set-ShapeTextFormat -Shape $FqdnLabel -Size "14 pt" -Case 1 -Style 1| Out-Null
	$FqdnLabel.Text = "$Name"
	$PoolNamesLayer.Add($FqdnLabel,1)

	$RoleLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 ($X1 + 2.5) -Y1 ($Y1 + 0.25) -Name "$($Name)RoleLabel" -Height "0.25 in" -Width 4 -Color $($Colors.White) -Transparency 100 -LineWeight "0 pt"
	Set-ShapeTextFormat -Shape $RoleLabel -Size "10 pt" -Case 1 -Style 2 | Out-Null
	$RoleLabel.Text = "Roles:`t$Roles"
	$RolesLayer.Add($RoleLabel,1)

	if($Members){
		$MembersLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 ($X1 + 2.5) -Y1 ($Y1 - 0.25) -Name "$($Name)MembersLabel" -Height "0.5 in" -Width 4 -Color $($Colors.White) -Transparency 100 -LineWeight "0 pt"
		Set-ShapeTextFormat -Shape $MembersLabel -Size "10 pt" -Case 1 | Out-Null
		$MembersLabel.Text = "Members:`t$Members"
		$PoolMembersLayer.Add($MembersLabel,1)
	}
	
	
	return $ShapeOnPage
}

function Add-ShapeToContainer {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$Shape,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$Container,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Resize = 0
	)
	
	$Result = $CurrentPage.Shapes.Item("$Container").ContainerProperties.AddMember($Shape, $Resize)

	return $Result
}

function Convert-ShapeToContainer {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$Shape
	)
	
	$Shape.AddNamedRow(242,"msvStructureType",0)
	$Shape.Cells("User.msvStructureType").Formula = """Container"""

	return $Shape
}

function Add-LineToPage {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$BeginX,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$BeginY,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$EndX,
		[Parameter(Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$EndY,
		[Parameter(Position = 4, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Name = $null,
		[Parameter(Position = 5, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Page = $CurrentPage,
		[Parameter(Position = 6, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Color,
		[Parameter(Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LineWeight = "0.75 pt",
		[Parameter(Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LinePattern = 1,
		[Parameter(Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LineType = 1,
		[Parameter(Position = 10, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$BeginArrow = 1,
		[Parameter(Position = 11, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$EndArrow = 1,
		[Parameter(Position = 12, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$BeginArrowSize = 1,
		[Parameter(Position = 13, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$EndArrowSize = 1
	)
	
	$LineOnPage = $Page.DrawLine(0,0,1,1)
	if($Name){$LineOnPage.Name = $Name}
	Update-Log "Adding line $Name to page."

	$LineOnPage.CellsU("BeginX").Formula = "$BeginX"
	$LineOnPage.CellsU("BeginY").Formula = "$BeginY"
	$LineOnPage.CellsU("EndX").Formula = "$EndX"
	$LineOnPage.CellsU("EndY").Formula = "$EndY"
	$LineOnPage.CellsU("LinePattern").Formula = "$LinePattern"
	$LineOnPage.CellsU("CompoundType").Formula = "$LineType"
	$LineOnPage.CellsU("LineWeight").Formula = "$LineWeight"
	$LineOnPage.CellsU("LineColor").Formula = "THEMEGUARD($Color)"
	$LineOnPage.CellsU("BeginArrow").Formula = "$BeginArrow"
	$LineOnPage.CellsU("EndArrow").Formula = "$EndArrow"
	$LineOnPage.CellsU("BeginArrowSize").Formula = "$BeginArrowSize"
	$LineOnPage.CellsU("EndArrowSize").Formula = "$EndArrowSize"
	$LineOnPage.Text = ""
	
	return $LineOnPage
}

function Add-RuleBraceToPage {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$LocX,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$LocY,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Name = $null,
		[Parameter(Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Page = $CurrentPage,
		[Parameter(Position = 4, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Height,
		[Parameter(Position = 5, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Width,
		[Parameter(Position = 6, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Direction,
		[Parameter(Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LineWeight = 0.75,
		[Parameter(Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$ResizeVertex = -0.125
	)
	
	$ShapeOnPage = $Page.Drop($VisioStencils.Brace, 10, 10)
	if($Name){$ShapeOnPage.Name = $Name}
	Update-Log "Adding rule brace $Name to page."

	$ShapeOnPage.CellsU("PinX").Formula = "$LocX"
	$ShapeOnPage.CellsU("PinY").Formula = "$LocY"
	$ShapeOnPage.CellsU("Height").Formula = "$Height"
	$ShapeOnPage.CellsU("Width").Formula = "$Width"
	
	if($Direction -eq "Left"){
		$ShapeOnPage.CellsU("Angle").Formula = "270 deg"
	} else {
		$ShapeOnPage.CellsU("Angle").Formula = "90 deg"
	}

	$ShapeOnPage.AddRow(7,0,0)
	$ShapeOnPage.CellsU("Connections.X1").Formula = "Width * 0.5"
	$ShapeOnPage.CellsU("Connections.Y1").Formula = "Height * -1"
	$ShapeOnPage.CellsU("LineWeight").Formula = "$LineWeight pt"
	$ShapeOnPage.CellsU("Controls.Y2").Formula = "$ResizeVertex in"
	$ShapeOnPage.Text = ""
	return $ShapeOnPage
}

function Add-ShapeToPage {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$X1,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Y1,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$X2 = $null,
		[Parameter(Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Y2 = $null,
		[Parameter(Position = 4, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$Shape,
		[Parameter(Position = 5, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Name = $null,
		[Parameter(Position = 6, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Page = $CurrentPage,
		[Parameter(Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Settings = $null,
		[Parameter(Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Height,
		[Parameter(Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Width,
		[Parameter(Position = 10, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Color,
		[Parameter(Position = 11, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LineColor = $null,
		[Parameter(Position = 12, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LineWeight = "0.75 pt",
		[Parameter(Position = 13, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LinePattern = 1,
		[Parameter(Position = 14, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Transparency = 0
	)
	
	$ShapeOnPage = $Page.Drop($($VisioStencils."$Shape"), 10, 10)
	if($Name){$ShapeOnPage.Name = $Name}
	
	Update-Log "Adding shape $Name to page."
	# If X2 is provided then this shape needs a beginning and ending location specified. This is for shapes like Arrows.
	if($X2){
		$ShapeOnPage.CellsU("BeginX").Formula = "$X1"
		$ShapeOnPage.CellsU("BeginY").Formula = "$Y1"
		$ShapeOnPage.CellsU("EndX").Formula = "$X2"
		$ShapeOnPage.CellsU("EndY").Formula = "$Y2"
	} else {
		$ShapeOnPage.CellsU("PinX").Formula = "$X1"
		$ShapeOnPage.CellsU("PinY").Formula = "$Y1"
	}
	if($Height){$ShapeOnPage.CellsU("Height").Formula = "$Height"}
	if($Height){$ShapeOnPage.CellsU("Width").Formula = "$Width"}
	if($Shape -eq "Square"){$ShapeOnPage.CellsU("Rounding").Formula = "0.1 in"}
	if($Color){
		$ShapeOnPage.CellsU("FillForegnd").Formula = "THEMEGUARD($Color)"
		$ShapeOnPage.CellsU("FillForegndTrans").Formula = "$Transparency%"
		$ShapeOnPage.CellsU("FillBkgnd").Formula = "0"
		$ShapeOnPage.CellsU("FillBkgndTrans").Formula = "$Transparency%"
		$ShapeOnPage.CellsU("FillPattern").Formula = "1"
		$ShapeOnPage.CellsU("LinePattern").Formula = "$LinePattern"
		$ShapeOnPage.CellsU("LineWeight").Formula = "$LineWeight"
		if(!$LineColor){$LineColor = $Color}
		$ShapeOnPage.CellsU("LineColor").Formula = "THEMEGUARD($LineColor)"
		$ShapeOnPage.CellsU("LineCap").Formula = "0"
		if($Shape -eq "Rectangle"){$ShapeOnPage.CellsU("Rounding").Formula = "0.1 in"}
	}
	
	if($Settings){
		$PropertyList = $Settings.GetEnumerator() | Select -ExpandProperty Name
		foreach($Property in $PropertyList){
			$ShapeOnPage.Cells("$Property").Formula = "$($Settings.$Property)"
		}
	}
	
	$ShapeOnPage.Text = ""
	return $ShapeOnPage
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

function New-ShapeConnectionPoint {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Shape,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[string[]]$ConnectionPoints
	)
	
	foreach($Location in $ConnectionPoints){
		# Add a new row to the ConnectionPoint section.
		$Shape.AddRow(7,0,0)
		$ShapeOnPage.CellsU("Connections.X1").Formula = "Width * 0.5"
		$ShapeOnPage.CellsU("Connections.Y1").Formula = "Height * -1"

	}
	$Shape.CellsU("Char.Font").Formula = "$Font"
	$Shape.CellsU("Char.Color").Formula = "THEMEGUARD($Color)"
	$Shape.CellsU("Char.Style").Formula = "$Style"
	$Shape.CellsU("Char.Case").Formula = "$Case"
	$Shape.CellsU("Para.HorzAlign").Formula = "$HAlign"
	$Shape.CellsU("VerticalAlign").Formula = "$VAlign"
	$Shape.CellsU("Para.IndLeft").Formula = "$LeftIndent"
	$Shape.CellsU("Para.IndRight").Formula = "$RightIndent"
	
	return $Shape
}

function New-VisioDiagram {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		[string] $DataFileName,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		$CsConfig,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[bool] $Visible
	)
	
	# Set the Visio document filename.
	$VisioDocFileName = $DataFileName.Replace(".xml",".vsdx")
	Update-Log "Creating diagram: $($VisioDocFileName)"
	
	# Create a new instance of Microsoft Visio to work with.
	Update-Log "Creating new instance of Visio to work with."
	$script:VisioApplication = New-Object -ComObject "Visio.Application"
	
	# Create a new blank document to work with and make the Visio application visible.
	Update-Log "Creating new Visio document."
	$VisioDocuments = $VisioApplication.Documents
	$VisioDocument = $VisioDocuments.Add("NETW_U.VST")
	$VisioPages = $VisioApplication.ActiveDocument.Pages
	Update-Log "Adding page(s) to Visio document."
	$VisioPage = $VisioPages.Item(1)
	$VisioDocument.PrintLandscape = $true
	$VisioPage.AutoSize = $false
	# Set Paper Size to Ledger (11" x 17")
	$VisioDocument.PaperSize = 3
	$VisioApplication.Visible = $Visible
	$FontIndex = $VisioDocument.Fonts | where{$_.Name -eq "Segoe UI"} | Select -ExpandProperty Index
	
	# Import and setup stencils for later use.
	Update-Log "Importing stencils..."
	[string]$VisioStencilPath = [System.Environment]::GetFolderPath('MyDocuments') + "\My Shapes"
	[string]$ServersStencil = "s_symbols_Servers_2014.vss"
	[string]$UsersStencil = "s_symbols_Users_2014.vss"
	[string]$CloudsStencil = "s_symbols_Clouds_2014.vss"
	[string]$ConceptsStencil = "s_symbols_Concepts_2014.vss"
	[string]$DevicesStencil = "s_symbols_Devices_2014.vss"
	$colServersStencil = $VisioApplication.Documents.OpenEx("$VisioStencilPath\$ServersStencil",4)
	$colUsersStencil = $VisioApplication.Documents.OpenEx("$VisioStencilPath\$UsersStencil",4)
	$colCloudsStencil = $VisioApplication.Documents.OpenEx("$VisioStencilPath\$CloudsStencil",4)
	$colConceptsStencil = $VisioApplication.Documents.OpenEx("$VisioStencilPath\$ConceptsStencil",4)
	$colDevicesStencil = $VisioApplication.Documents.OpenEx("$VisioStencilPath\$DevicesStencil",4)
	$colConnectorsStencils = $VisioApplication.Documents.OpenEx("CONNEC_U.VSSX",4)
	$colCalloutStencils = $VisioApplication.Documents.OpenEx("CALOUT_U.VSSX",4)
	$colBasicStencils = $VisioApplication.Documents.OpenEx("BASIC_U.VSSX",4)
	
	# Create hashtable for stencils and shapes.
	$VisioStencils = @{
		"ApplicationServer" = $colServersStencil.Masters.Item("Trusted Application Server")
		"DatabaseServer" = $colServersStencil.Masters.Item("Database Server")
		"Director" = $colServersStencil.Masters.Item("Skype for Business Director")
		"EdgeServer" = $colServersStencil.Masters.Item("Skype for Business Edge Server")
		"FEPool" = $colServersStencil.Masters.Item("Skype for Business Front-End Pool")
		"FEServer" = $colServersStencil.Masters.Item("Skype for Business Front-End Server")
		"FileStore" = $colServersStencil.Masters.Item("File Server")
		"Firewall" = $colConceptsStencil.Masters.Item("Firewall")
		"LoadBalancer" = $colDevicesStencil.Masters.Item("Load Balancer")
		"MediationServer" = $colServersStencil.Masters.Item("Skype for Business Mediation Server")
		"MonitoringServer" = $colServersStencil.Masters.Item("Skype for Business Monitoring Server")
		"ReverseProxy" = $colServersStencil.Masters.Item("Reverse Proxy")
		"SBASBS" = $colServersStencil.Masters.Item("Survivable Branch Server")
		"Server" = $colServersStencil.Masters.Item("Server, Generic")
		"SFBServer" = $colServersStencil.Masters.Item("Skype for Business Server")
		"IPGateway" = $colDevicesStencil.Masters.Item("IP Gateway")
		"MobileUser" = $colUsersStencil.Masters.Item("Mobile User")
		"SFBUser" = $colUsersStencil.Masters.Item("Skype for Business User")
		"SkypeUser" = $colUsersStencil.Masters.Item("Skype Commercial User")
		"Arrow" = $colConnectorsStencils.Masters.Item("1-D Single")
		"DoubleArrow" = $colConnectorsStencils.Masters.Item("1-D double")
		"Connector" = $colConnectorsStencils.Masters.Item("Dynamic connector")
		"Brace" = $colCalloutStencils.Masters.Item("Side Brace")
		"RoundRectangle" = $colBasicStencils.Masters.Item("Rounded Rectangle")
		"Rectangle" = $colBasicStencils.Masters.Item("Rectangle")
	}
	
	for($CurrentSiteNumber = 1; $CurrentSiteNumber -le $CsConfig.Topology.Object.Sites.Count; $CurrentSiteNumber++ ){
		# Select page for current site.
		$CurrentPage = $VisioPages.Item($CurrentSiteNumber)
		
		# Build Layers for current page.
		$BackgroundLayer = $CurrentPage.Layers.Add("Border-Legend")
		$ConnectorsLayer = $CurrentPage.Layers.Add("Connectors")
		$DeploymentDetailsLayer = $CurrentPage.Layers.Add("Deployment Details")
		$FirewallRulesLayer = $CurrentPage.Layers.Add("Firewall Rules")
		$PoolNamesLayer = $CurrentPage.Layers.Add("Pool Names")
		$PoolMembersLayer = $CurrentPage.Layers.Add("Pool Members")
		$RolesLayer = $CurrentPage.Layers.Add("Roles")
		$SiteBGLayer = $CurrentPage.Layers.Add("Network Segment Backgrounds")
		$ServersLayer = $CurrentPage.Layers.Add("Servers")

		# Make sure there are enough pages for each site.
		if ($VisioPages.Count -lt $CurrentSiteNumber){$CurrentPage = $VisioPages.Add()}
		$CurrentPage.PageSheet.CellsU("PageWidth").Formula= "33.25 in"
		$CurrentPage.PageSheet.CellsU("PageHeight").Formula= "21.25 in"
		
		$Site = $CsConfig.Topology.Object.Sites[$CurrentSiteNumber - 1]
		$CurrentPage.Name = $Site.Name
		New-VisioSitePage -Site $Site -CurrentPage $CurrentPage
	}
	Update-Status "Done creating diagram."
	Update-Status "Saving changes."
	$VisioDocument.SaveAs($VisioDocFileName) | Out-Null
	Update-Status "Done!"
}

function New-VisioPageTemplate{
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$CurrentPage
	)
	# Black page border
	$DocumentOutline = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.5" -Y1 "ThePage!PageHeight * 0.5" -Name "DocumentOutline" -Height "ThePage!PageHeight - 0.5" -Width "ThePage!PageWidth - 0.5" -Color $($Colors.Black) -LineWeight "2.5 pt" -Transparency 100
	$BackgroundLayer.Add($DocumentOutline,1)
	
	# Title Box
	$TitleBox = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "((ThePage!PageWidth - 0.5) * 0.25) + 0.25" -Y1 "0.5 in" -Name "TitleBox" -Height "0.5 in" -Width "(ThePage!PageWidth - 0.5) * 0.5" -Color $($Colors.DarkGray) -Transparency 0
	Set-ShapeTextFormat -Shape $TitleBox -Size "24 pt" -Color "$($Colors.White)" -Style 1 -VAlign 1 -LeftIndent "0.25 in"| Out-Null
	$TitleBox.Text = "Skype for Business Environment Diagram"
	$BackgroundLayer.Add($TitleBox,1)
	
	# Second box
	$TimeStampBox = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "((ThePage!PageWidth - 0.5) * 0.65) + 0.25" -Y1 "0.5 in" -Name "TimeStampBox" -Height "0.5 in" -Width "(ThePage!PageWidth - 0.5) * 0.30" -Color $($Colors.DarkGray) -Transparency 40
	Set-ShapeTextFormat -Shape $TimeStampBox -Size "18 pt" -Style 1 -HAlign 1 -VAlign 1 | Out-Null
	$TimeStampBox.Text = "Data Collected: $($CsConfig.TimeStamp)"
	$BackgroundLayer.Add($TimeStampBox,1)
	
	# Third box
	$EMBox = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "((ThePage!PageWidth - 0.5) * 0.9) + 0.25" -Y1 "0.5 in" -Name "EMBox" -Height "0.5 in" -Width "(ThePage!PageWidth - 0.5) * 0.20" -Color $($Colors.DarkGray) -Transparency 80
	Set-ShapeTextFormat -Shape $EMBox -Size "18 pt" -Style 1 -HAlign 1 -VAlign 1 | Out-Null
	$EMBox.Text = "EmptyMessage.com"
	$BackgroundLayer.Add($EMBox,1)
	
	# Outside User / External Network Background
	$ExternalDmzBox = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.125" -Y1 "ThePage!PageHeight * 0.45" -Name "ExternalDmzBox" -Height "ThePage!PageHeight * 0.75" -Width "ThePage!PageWidth * 0.225" -Color $($Colors.Red) -LinePattern 23 -LineWeight "0.044 pt*ThePage!PageWidth" -Transparency 90
	Set-ShapeTextFormat -Shape $ExternalDmzBox -Size "1 pt*ThePage!PageWidth" -Style 1 -Case 1 -HAlign 2 -VAlign 2 -RightIndent "0.25 in" | Out-Null
	$ExternalDmzBox.Text = "Outside Users"
	$SiteBGLayer.Add($ExternalDmzBox,1)
	Convert-ShapeToContainer -Shape $ExternalDmzBox | Out-Null
	
	# Internal DMZ Network Background
	$InternalDmzBox = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.375" -Y1 "ThePage!PageHeight * 0.45" -Name "InternalDmzBox" -Height "ThePage!PageHeight * 0.75" -Width "ThePage!PageWidth * 0.225" -Color $($Colors.Orange) -LinePattern 23 -LineWeight "0.044 pt*ThePage!PageWidth" -Transparency 90
	Set-ShapeTextFormat -Shape $InternalDmzBox -Size "1 pt*ThePage!PageWidth" -Style 1 -Case 1 -HAlign 2 -VAlign 2 -RightIndent "0.25 in" | Out-Null
	$InternalDmzBox.Text = "DMZ Servers"
	$SiteBGLayer.Add($InternalDmzBox,1)
	Convert-ShapeToContainer -Shape $InternalDmzBox | Out-Null
	
	# Internal Network Background
	$InternalNetworkBox = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.75" -Y1 "ThePage!PageHeight * 0.45" -Name "InternalNetworkBox" -Height "ThePage!PageHeight * 0.75" -Width "ThePage!PageWidth * 0.475" -Color $($Colors.Green) -LinePattern 23 -LineWeight "0.044 pt*ThePage!PageWidth" -Transparency 90
	Set-ShapeTextFormat -Shape $InternalNetworkBox -Size "1 pt*ThePage!PageWidth" -Style 1 -Case 1 -HAlign 2 -VAlign 2 -RightIndent "0.25 in" | Out-Null
	$InternalNetworkBox.Text = "Internal Servers"
	$SiteBGLayer.Add($InternalNetworkBox,1)
	Convert-ShapeToContainer -Shape $InternalNetworkBox | Out-Null
	
	# External Firewall Divider Arrow
	$FirewallArrow1 = Add-LineToPage -Page $CurrentPage -Name "FirewallArrow1" -BeginX "ThePage!PageWidth * 0.25" -BeginY "ThePage!PageHeight * 0.075" -EndX "ThePage!PageWidth * 0.25" -EndY "ThePage!PageHeight * 0.90" -LineWeight "0.147 pt*ThePage!PageWidth" -Color $($Colors.Red) -BeginArrow "13" -EndArrow "13"
	$BackgroundLayer.Add($FirewallArrow1,1)
	
	# Firewall Stencil
	$ExternalFirewall = Add-ShapeToPage -Shape "Firewall" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.25" -Y1 "ThePage!PageHeight * 0.85" -Name "ExternalFirewall"
	$BackgroundLayer.Add($ExternalFirewall,1)
	
	# External Firewall Label
	$ExternalFirewallLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.25" -Y1 "ThePage!PageHeight * 0.9125" -Name "ExternalFirewallLabel" -Height "ThePage!PageHeight * 0.025" -Width "ThePage!PageWidth * 0.15" -Color $($Colors.Gray) -LineWeight "0.044 pt*ThePage!PageWidth" -Transparency 90
	Set-ShapeTextFormat -Shape $ExternalFirewallLabel -Size "1 pt*ThePage!PageWidth" -Style 1 -Case 1 -HAlign 1 -VAlign 1| Out-Null
	$ExternalFirewallLabel.Text = "External Firewall"
	$BackgroundLayer.Add($ExternalFirewallLabel,1)
	
	# Internal Firewall Divider Arrow
	$FirewallArrow2 = Add-LineToPage -Page $CurrentPage -Name "FirewallArrow2" -BeginX "ThePage!PageWidth * 0.5" -BeginY "ThePage!PageHeight * 0.075" -EndX "ThePage!PageWidth * 0.5" -EndY "ThePage!PageHeight * 0.90" -LineWeight "0.147 pt*ThePage!PageWidth" -Color $($Colors.Red) -BeginArrow "13" -EndArrow "13"
	$BackgroundLayer.Add($FirewallArrow2,1)
	
	# Firewall Stencil
	$InternalFirewall = Add-ShapeToPage -Shape "Firewall" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.5" -Y1 "ThePage!PageHeight * 0.85" -Name "InternalFirewall"
	$BackgroundLayer.Add($InternalFirewall,1)
	
	# Internal Firewall Label
	$InternalFirewallLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.5" -Y1 "ThePage!PageHeight * 0.9125" -Name "InternalFirewallLabel" -Height "ThePage!PageHeight * 0.025" -Width "ThePage!PageWidth * 0.15" -Color $($Colors.Gray) -LineWeight "0.044 pt*ThePage!PageWidth" -Transparency 90
	Set-ShapeTextFormat -Shape $InternalFirewallLabel -Size "1 pt*ThePage!PageWidth" -Style 1 -Case 1 -HAlign 1 -VAlign 1| Out-Null
	$InternalFirewallLabel.Text = "Internal Firewall"
	$BackgroundLayer.Add($InternalFirewallLabel,1)
	
	# Legend
	$LegendLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.7875" -Y1 "ThePage!PageHeight * 0.9" -Name "LegendLabel" -Height "ThePage!PageHeight * 0.125" -Width "ThePage!PageWidth * 0.4" -Color $($Colors.Gray) -LineWeight "2.5 pt" -Transparency 90
	Set-ShapeTextFormat -Shape $LegendLabel -Size "1 pt * ThePage!PageWidth" -Style 1 -LeftIndent "0.14 in * ThePage!PageWidth"| Out-Null
	$LegendLabel.Text = "FE Pool to SQL Backend`nMediation To Gateway Trunk`nFE Pool to Trusted Application"
	$BackgroundLayer.Add($LegendLabel,1)
	Convert-ShapeToContainer -Shape $LegendLabel | Out-Null
	
	# Legend Arrows
	$FESQLArrow = Add-LineToPage -Page $CurrentPage -Name "FESQLArrow" -BeginX "ThePage!PageWidth * 0.62" -BeginY "ThePage!PageHeight * 0.945" -EndX "ThePage!PageWidth * 0.72" -EndY "ThePage!PageHeight * 0.945" -LineWeight "0.147 pt*ThePage!PageWidth" -Color $($Colors.Red) -BeginArrow "5" -EndArrow "5" -LineType 1
	$BackgroundLayer.Add($FESQLArrow,1)
	
	$MedGWArrow = Add-LineToPage -Page $CurrentPage -Name "MedGWArrow" -BeginX "ThePage!PageWidth * 0.62" -BeginY "ThePage!PageHeight * 0.9175" -EndX "ThePage!PageWidth * 0.72" -EndY "ThePage!PageHeight * 0.9175" -LineWeight "0.147 pt*ThePage!PageWidth" -Color $($Colors.Orange) -BeginArrow "5" -EndArrow "5" -LineType 4
	$BackgroundLayer.Add($MedGWArrow,1)
	
	$FEAppArrow = Add-LineToPage -Page $CurrentPage -Name "FEAppArrow" -BeginX "ThePage!PageWidth * 0.62" -BeginY "ThePage!PageHeight * 0.89" -EndX "ThePage!PageWidth * 0.72" -EndY "ThePage!PageHeight * 0.89" -LineWeight "0.147 pt*ThePage!PageWidth" -Color $($Colors.Green) -BeginArrow "5" -EndArrow "5" -LineType 0
	$BackgroundLayer.Add($FEAppArrow,1)
}

function New-VisioSitePage{
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Site,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$CurrentPage
	)
	
	New-VisioPageTemplate -CurrentPage $CurrentPage
	$XmlTopology = [xml]$CsConfig.Topology.XML

	# Get number of Edge servers in site.
	$SiteEdgeServerIds = $Site.Clusters| where {$_.IsOnEdge} | select -ExpandProperty Machines
	$SiteEdgeServers = @()
	foreach ($ServerId in $SiteEdgeServerIds){$SiteEdgeServers += $CsConfig.Topology.Object.Machines | where{$_.MachineId -contains $ServerId}}
	
	# Enumerate Edge servers and build section for each one.
	for ($i = 0; $i -lt $SiteEdgeServerIds.Count; $i++){
		$EdgeHeight = (-1 + (($i + 1) * 6.5))
		
		#Add Edge Server shape to page.
		$EdgeServerShape = Add-ShapeToPage -Shape "EdgeServer" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.375" -Y1 "$($EdgeHeight - 0.75) in" -Name "EdgeServer$i"
		$ServersLayer.Add($EdgeServerShape,1)
		Add-ShapeToContainer -Shape $EdgeServerShape -Container "InternalDmzBox" -Resize 0
		
		$TopologyMachine = $CsConfig.Topology.Object.Machines | where{$_.MachineId -match $SiteEdgeServerIds[$i]}
		$Machine = $CsConfig.EnvironmentData.ServerData."$($TopologyMachine.Fqdn)"
		$XMLCluster = $script:XMLTopology.Topology.Clusters.Cluster | where{$_.Fqdn -match $TopologyMachine.Cluster.Fqdn}
		$XMLMachine = $XMLCluster.Machine | where{$_.fqdn -match $TopologyMachine.Fqdn}
		$ServiceId = $TopologyMachine.Cluster.InstalledServices | where{$_ -match "Edge"}
		$ServiceIdSplit = $ServiceId -split "-"
		$EdgeService = $XMLTopology.Topology.Services.Service | where{($_.ServiceId.SiteId -match $ServiceIdSplit[0]) -and ($_.ServiceId.RoleName -match $ServiceIdSplit[1]) -and ($_.ServiceId.Instance -match $ServiceIdSplit[2])}
		$AEFqdn = $EdgeService.Ports.Port | ? {($_.Owner -match "Access") -and ($_.InterfaceSide -match "External")} | Select -Unique -ExpandProperty ConfiguredFqdn
		$WCFqdn = $EdgeService.Ports.Port | ? {($_.Owner -match "Data") -and ($_.InterfaceSide -match "External")} | Select -Unique -ExpandProperty ConfiguredFqdn
		$AVFqdn = $EdgeService.Ports.Port | ? {($_.Owner -match "Media") -and ($_.InterfaceSide -match "External")} | Select -Unique -ExpandProperty ConfiguredFqdn
		$InternalIPAddress = $XMLMachine.NetInterface | where{$_.InterfaceSide -eq "Internal"} | Select -ExpandProperty IPAddress
		[string[]]$ExternalIPAddresses = $XMLMachine.NetInterface | where{$_.InterfaceSide -eq "External"} | sort InterfaceNumber | Select -ExpandProperty IPAddress

		
		# Create External firewall rule arrows.
		for ($CurrentRuleNumber = 0; $CurrentRuleNumber -lt ($ExternalEdgeFirewallRules -split "`n").Count; $CurrentRuleNumber++){
			$LocY = ($EdgeHeight + 2.75) - (($CurrentRuleNumber) * 0.5)
			$Rule = ($ExternalEdgeFirewallRules -split "`n")[$CurrentRuleNumber]
			$RuleDetails = $Rule -split ","
			switch ($RuleDetails[1]) { 
				"Inbound" {
					$BeginX = "(ThePage!PageWidth * 0.25) - 1"
					$EndX = "(ThePage!PageWidth * 0.25) + 1"
					$Stencil = "Arrow"
				}
				"Outbound" {
					$BeginX = "(ThePage!PageWidth * 0.25) + 1"
					$EndX = "(ThePage!PageWidth * 0.25) - 1"
					$Stencil = "Arrow"
				}
				"Both" {
					$BeginX = "(ThePage!PageWidth * 0.25) - 1"
					$EndX = "(ThePage!PageWidth * 0.25) + 1"
					$Stencil = "DoubleArrow"
				}
			}
			$FirewallRuleArrow = Add-ShapeToPage -Shape "$Stencil" -Page $CurrentPage -X1 $BeginX -Y1 $LocY -X2 $EndX -Y2 $LocY -Name "ExternalEdge$($i)Rule$($CurrentRuleNumber + 1)" -Height "0.5 in" -Color $($Colors.White) -Transparency 20 -LineColor $($Colors.Red)
			Set-ShapeTextFormat -Shape $FirewallRuleArrow -Size "8 pt" -Style 1 -HAlign 1 -VAlign 1 | Out-Null
			$FirewallRuleArrow.Text = "$($RuleDetails[0])"
			$FirewallRulesLayer.Add($FirewallRuleArrow,1)
			Add-ShapeToContainer -Shape $FirewallRuleArrow -Container "InternalDmzBox" -Resize 0
			if($RuleDetails[1] -eq "Both"){
				$FirewallRuleArrow.CellsU("Scratch.A1").Formula = "0.5"
				$FirewallRuleArrow.CellsU("Scratch.B1").Formula = "0.25"				
			} else {
				$FirewallRuleArrow.CellsU("Scratch.X2").Formula = "0.5"
				$FirewallRuleArrow.CellsU("Scratch.Y2").Formula = "0.25"				
			}
		}
		# Create braces for rule groupings.
		$AELeftBrace = Add-RuleBraceToPage -LocX "(ThePage!PageWidth * 0.25) - 1" -LocY "$($EdgeHeight + 1.25) in" -Name "AELeftBrace$i" -Height "0.125 in" -Width "3.5 in" -LineWeight "1.25" -Direction "Left"
		$AERightBrace = Add-RuleBraceToPage -LocX "(ThePage!PageWidth * 0.25) + 1" -LocY "$($EdgeHeight + 1.25) in" -Name "AERightBrace$i" -Height "0.125 in" -Width "3.5 in" -LineWeight "1.25" -Direction "Right"
		$AVLeftBrace = Add-RuleBraceToPage -LocX "(ThePage!PageWidth * 0.25) - 1" -LocY "$($EdgeHeight - 2) in" -Name "AVLeftBrace$i" -Height "0.125 in" -Width "2 in" -LineWeight "1.25" -Direction "Left"
		$AVRightBrace = Add-RuleBraceToPage -LocX "(ThePage!PageWidth * 0.25) + 1" -LocY "$($EdgeHeight - 2) in" -Name "AVRightBrace$i" -Height "0.125 in" -Width "2 in" -LineWeight "1.25" -Direction "Right"
		
		#Access Edge label
		$AEDetailsLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "(ThePage!PageWidth * 0.25) - 3.75" -Y1 "$($EdgeHeight + 1.375) in" -Name "AEDetailsLabel$($i)" -LineColor $($Colors.Gray) -Color $($Colors.White) -Height 3.25 -Width 5 -Transparency 20
		$AEDetailsText = @()
		$AEDetailsText += "Access Edge"
		$AEDetailsText += "FQDN:`t$AEFqdn"
		[string[]]$SipFqdns = $CsConfig.EnvironmentData.ExternalDNS.GetEnumerator() | ? {($_.Name -match "sip.") -and ($_.Name -notmatch "_")} | Select -ExpandProperty Name
		foreach($Fqdn in $SipFqdns){
			$AEDetailsText += "FQDN:`t$Fqdn"
		}
		if($ExternalIPAddresses.Count -gt 1){$AEDetailsText += "IP:`t$ExternalIPAddresses[0]"} else {$AEDetailsText += "IP:`t$ExternalIPAddresses"}
		Set-ShapeTextFormat -Shape $AEDetailsLabel -Size "14 pt" -Style 0 -Case 1 -Valign 1| Out-Null
		$AEDetailsLabel.Text = $AEDetailsText -join "`n"
		Add-ShapeToContainer -Shape $AEDetailsLabel -Container "InternalDmzBox" -Resize 0

		#Web Conferencing Edge label
		$WCDetailsLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "(ThePage!PageWidth * 0.25) - 3.75" -Y1 "$($EdgeHeight - 0.75) in" -Name "WCDetailsLabel$($i)" -LineColor $($Colors.Gray) -Color $($Colors.White) -Height 0.75 -Width 5 -Transparency 20
		if($ExternalIPAddresses.Count -gt 1){$WCIP = $ExternalIPAddresses[1]} else {$WCIP = $ExternalIPAddresses}
		Set-ShapeTextFormat -Shape $WCDetailsLabel -Size "14 pt" -Style 0 -Case 1 -Valign 1| Out-Null
		$WCDetailsLabel.Text = "Web Conferencing Edge`nFQDN:  $WCFqdn`nIP:`t$WCIP"
		Add-ShapeToContainer -Shape $WCDetailsLabel -Container "InternalDmzBox" -Resize 0

		#Audio Video Edge label
		$AVDetailsLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "(ThePage!PageWidth * 0.25) - 3.75" -Y1 "$($EdgeHeight - 2.125) in" -Name "AVDetailsLabel$($i)" -LineColor $($Colors.Gray) -Color $($Colors.White) -Height 1.75 -Width 5 -Transparency 20
		if($ExternalIPAddresses.Count -gt 1){$AVIP = $ExternalIPAddresses[2]} else {$AVIP = $ExternalIPAddresses}
		Set-ShapeTextFormat -Shape $AVDetailsLabel -Size "14 pt" -Style 0 -Case 1 -Valign 1| Out-Null
		$AVDetailsLabel.Text = "Audio Video Edge`nFQDN:  $WCFqdn`nIP:`t$AVIP"
		Add-ShapeToContainer -Shape $AVDetailsLabel -Container "InternalDmzBox" -Resize 0

		# Create connectors from rules/braces to Edge server shape.
		$AEConnector = Add-ConnectorToPage -Page $CurrentPage  -Name "ConnectorAE$i" -BeginX 1 -BeginY 1 -EndX 2 -EndY 2 -LineWeight "2 pt" -Color $($Colors.DarkGray) -LinePattern 2
		$AEConnector.Cells("BeginX").GlueTo($CurrentPage.Shapes.Item("AERightBrace$i").Cells("Connections.X1"))
		$AEConnector.Cells("EndX").GlueTo($EdgeServerShape.CellsU("Connections.X3"))
		Set-ShapeTextFormat -Shape $AEConnector -Size "12 pt" -Style 1 -Case 1 -HAlign 1 -VAlign 1 | Out-Null
		$AEConnector.Cells("TxtAngle").Formula = "-40 deg"
		if($ExternalIPAddresses.Count -gt 1){$AEConnector.Text = "$ExternalIPAddresses[0]"} else {$AEConnector.Text = "$ExternalIPAddresses"}
		Add-ShapeToContainer -Shape $AEConnector -Container "InternalDmzBox" -Resize 0
		
		$WCConnector = Add-ConnectorToPage -Page $CurrentPage  -Name "ConnectorAV$i" -BeginX 1 -BeginY 1 -EndX 2 -EndY 2 -LineWeight "2 pt" -Color $($Colors.DarkGray) -LinePattern 2
		$WCConnector.Cells("BeginX").GlueTo($CurrentPage.Shapes.Item("ExternalEdge$($i)Rule8").Cells("Geometry1.X5"))
		$WCConnector.Cells("EndX").GlueTo($EdgeServerShape.CellsU("Connections.X3"))
		Set-ShapeTextFormat -Shape $WCConnector -Size "12 pt" -Style 1 -Case 1 -HAlign 1 -VAlign 1 | Out-Null
		if($ExternalIPAddresses.Count -gt 1){$WCConnector.Text = "$ExternalIPAddresses[1]"} else {$WCConnector.Text = "$ExternalIPAddresses"}
		Add-ShapeToContainer -Shape $WCConnector -Container "InternalDmzBox" -Resize 0
		
		$AVConnector = Add-ConnectorToPage -Page $CurrentPage  -Name "ConnectorAV$i" -BeginX 1 -BeginY 1 -EndX 2 -EndY 2 -LineWeight "2 pt" -Color $($Colors.DarkGray) -LinePattern 2
		$AVConnector.Cells("BeginX").GlueTo($CurrentPage.Shapes.Item("AVRightBrace$i").Cells("Connections.X1"))
		$AVConnector.Cells("EndX").GlueTo($EdgeServerShape.CellsU("Connections.X3"))
		Set-ShapeTextFormat -Shape $AVConnector -Size "12 pt" -Style 1 -Case 1 -HAlign 1 -VAlign 1 | Out-Null
		$AVConnector.Cells("TxtAngle").Formula = "29 deg"
		if($ExternalIPAddresses.Count -gt 1){$AVConnector.Text = "$ExternalIPAddresses[2]"} else {$AVConnector.Text = "$ExternalIPAddresses"}
		Add-ShapeToContainer -Shape $AVConnector -Container "InternalDmzBox" -Resize 0
		
		# Create Internal firewall rule arrows.
		for ($CurrentRuleNumber = 0; $CurrentRuleNumber -lt ($InternalEdgeFirewallRules -split "`n").Count; $CurrentRuleNumber++){
			$LocY = ($EdgeHeight + 2.25) - (($CurrentRuleNumber) * 0.5)
			$Rule = ($InternalEdgeFirewallRules -split "`n")[$CurrentRuleNumber]
			$RuleDetails = $Rule -split ","
			switch ($RuleDetails[1]) { 
				"Inbound" {
					$BeginX = "(ThePage!PageWidth * 0.5) + 1"
					$EndX = "(ThePage!PageWidth * 0.5) - 1"
					$Stencil = "Arrow"
				}
				"Outbound" {
					$BeginX = "(ThePage!PageWidth * 0.5) - 1"
					$EndX = "(ThePage!PageWidth * 0.5) + 1"
					$Stencil = "Arrow"
				}
				"Both" {
					$BeginX = "(ThePage!PageWidth * 0.5) - 1"
					$EndX = "(ThePage!PageWidth * 0.5) + 1"
					$Stencil = "DoubleArrow"
				}
			}
			$FirewallRuleArrow = Add-ShapeToPage -Shape "$Stencil" -Page $CurrentPage -X1 $BeginX -Y1 $LocY -X2 $EndX -Y2 $LocY -Name "InternalEdge$($i)Rule$($CurrentRuleNumber + 1)" -Height "0.5 in" -Color $($Colors.White) -Transparency 20 -LineColor $($Colors.Red)
			Set-ShapeTextFormat -Shape $FirewallRuleArrow -Size "8 pt" -Style 1 -HAlign 1 -VAlign 1 | Out-Null
			$FirewallRuleArrow.Text = "$($RuleDetails[0])"
			$FirewallRulesLayer.Add($FirewallRuleArrow,1)
			Add-ShapeToContainer -Shape $FirewallRuleArrow -Container "InternalDmzBox" -Resize 0
			if($RuleDetails[1] -eq "Both"){
				$FirewallRuleArrow.CellsU("Scratch.A1").Formula = "0.5"
				$FirewallRuleArrow.CellsU("Scratch.B1").Formula = "0.25"				
			} else {
				$FirewallRuleArrow.CellsU("Scratch.X2").Formula = "0.5"
				$FirewallRuleArrow.CellsU("Scratch.Y2").Formula = "0.25"				
			}
		}
		# Create braces for rule groupings.
		$EdgeInternalLeftBrace = Add-RuleBraceToPage -LocX "(ThePage!PageWidth * 0.5) - 1" -LocY "$($EdgeHeight - 0) in" -Name "EdgeInternal$($i)RulesBrace$i" -Height "0.125 in" -Width "5 in" -LineWeight "1.25" -Direction "Left"
		$EdgeFERightBrace = Add-RuleBraceToPage -LocX "(ThePage!PageWidth * 0.5) + 1" -LocY "$($EdgeHeight + 1) in" -Name "EdgeFE$($i)RulesBrace$i" -Height "0.125 in" -Width "3 in" -LineWeight "1.25" -Direction "Right"
		$EdgeAnyRightBrace = Add-RuleBraceToPage -LocX "(ThePage!PageWidth * 0.5) + 1" -LocY "$($EdgeHeight - 1.5) in" -Name "EdgeAny$($i)RulesBrace$i" -Height "0.125 in" -Width "2 in" -LineWeight "1.25" -Direction "Right"
		
		# Create connectors from rules/braces to Edge server shape.
		$EdgeInternalConnector = Add-ConnectorToPage -Page $CurrentPage  -Name "ConnectorEdgeInternal$i" -BeginX 1 -BeginY 1 -EndX 2 -EndY 2 -LineWeight "2 pt" -Color $($Colors.DarkGray) -LinePattern 2
		$CurrentPage.Shapes.Item("EdgeInternal$($i)RulesBrace$i").Cells("Controls.Row_1").Formula = "(Width/2)-0.75 in"
		$CurrentPage.Shapes.Item("EdgeInternal$($i)RulesBrace$i").Cells("Connections.X1").Formula = "(Width/2)-0.75 in"
		$EdgeInternalConnector.Cells("BeginX").GlueTo($CurrentPage.Shapes.Item("EdgeInternal$($i)RulesBrace$i").Cells("Connections.X1"))
		$EdgeInternalConnector.Cells("EndX").GlueTo($EdgeServerShape.CellsU("Connections.X2"))
		Set-ShapeTextFormat -Shape $EdgeInternalConnector -Size "10 pt" -Style 1 -Case 1 -HAlign 1 -VAlign 1 | Out-Null
		$EdgeInternalConnector.Text = "$InternalIPAddress"
		Add-ShapeToContainer -Shape $EdgeInternalConnector -Container "InternalDmzBox" -Resize 0
		
		# Edge Server label
		$EdgeLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.375" -Y1 "$($EdgeHeight - 1.5) in" -Name "Edge$($i)Label" -LineColor $($Colors.Gray) -Color $($Colors.White) -Height "0.25 in" -Width 1.5 -Transparency 20
		Set-ShapeTextFormat -Shape $EdgeLabel -Size "12 pt" -Style 1 -Case 1 -HAlign 1 -Valign 1| Out-Null
		$EdgeLabel.Text = "Edge Server"
		Add-ShapeToContainer -Shape $EdgeLabel -Container "InternalDmzBox" -Resize 0
		$EdgeDetails = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.375" -Y1 "$($EdgeHeight - 2.5) in" -Name "Edge$($i)Details" -LineColor $($Colors.Gray) -Color $($Colors.White) -Height 0.5 -Width 3.75 -Transparency 20
		Set-ShapeTextFormat -Shape $EdgeDetails -Size "10 pt" -Style 0 -Case 1| Out-Null
		$EdgeDetails.Text = "Pool FQDN:`t$($TopologyMachine.Cluster.Fqdn)`nServer FQDN:`t$($TopologyMachine.Fqdn)"
		Add-ShapeToContainer -Shape $EdgeDetails -Container "InternalDmzBox" -Resize 0
		
		#Internal Front-End Server Rules label
		$FERulesLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "(ThePage!PageWidth * 0.5) + 2.375" -Y1 "$($EdgeHeight + 1) in" -Name "FERulesLabel$($i)" -LineColor $($Colors.Gray) -Color $($Colors.White) -Height "0.75 in" -Width 2.25 -Transparency 75
		Set-ShapeTextFormat -Shape $FERulesLabel -Size "14 pt" -Style 1 -Case 1 -HAlign 1 -Valign 1| Out-Null
		$FERulesLabel.Text = "All Front-End Pool and Front-End Server Addresses"
		Add-ShapeToContainer -Shape $FERulesLabel -Container "InternalNetworkBox" -Resize 0
		#Internal Clients and Addresses Rules label
		$ClientRulesLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "(ThePage!PageWidth * 0.5) + 2.375" -Y1 "$($EdgeHeight - 1.5) in" -Name "ClientRulesLabel$($i)" -LineColor $($Colors.Gray) -Color $($Colors.White) -Height "0.75 in" -Width 2.25 -Transparency 75
		Set-ShapeTextFormat -Shape $ClientRulesLabel -Size "14 pt" -Style 1 -Case 1 -HAlign 1 -Valign 1| Out-Null
		$ClientRulesLabel.Text = "All Internal Address Space"
		Add-ShapeToContainer -Shape $ClientRulesLabel -Container "InternalNetworkBox" -Resize 0
	}
	

	$InternalPools = $Site.Clusters| where {(!$_.IsOnEdge) -and ($_.Machines.Count -gt 1) -and ($_.InstalledServices -notmatch "PstnGateway") -and ($_.InstalledServices -notmatch "ExternalServer")}
	$InternalServers = $Site.Clusters| where {(!$_.IsOnEdge) -and ($_.Machines.Count -eq 1) -and ($_.InstalledServices -notmatch "PstnGateway") -and ($_.InstalledServices -notmatch "ExternalServer")}
	$ExternalPools = $Site.Clusters| where {($_.InstalledServices -match "ExternalServer") -and ($_.Machines.Count -gt 1)}
	$ExternalServers = $Site.Clusters| where {($_.InstalledServices -match "ExternalServer") -and ($_.Machines.Count -eq 1)}
	$PstnGateways = $Site.Clusters| where {($_.InstalledServices -match "PstnGateway")}
	$SqlInstances = $Site.Clusters| where {($_.SqlInstances)}
	
	$InternalObjectCount = $InternalPools + $InternalServers | measure | select -ExpandProperty Count
	$ExternalObjectCount = $ExternalPools + $ExternalServers | measure | select -ExpandProperty Count
	$PstnGatewayCount = $PstnGateways | measure | select -ExpandProperty Count
	$SqlInstanceCount = $SqlInstances | measure | select -ExpandProperty Count
	
	# Skype for Business Server label
	$SfBLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.70" -Y1 "ThePage!PageHeight * 0.625" -Name "SfBLabel" -LineColor $($Colors.Green) -Color $($Colors.White) -Height "ThePage!PageHeight * 0.375" -Width "ThePage!PageWidth * 0.1725" -Transparency 30
	Set-ShapeTextFormat -Shape $SfBLabel -Size "18 pt" -Style 1 -Case 1 -HAlign 1 | Out-Null
	$SfBLabel.Text = "Skype for Business Pools and Servers"
	Add-ShapeToContainer -Shape $SfBLabel -Container "InternalNetworkBox" -Resize 0
	Convert-ShapeToContainer -Shape $SfBLabel | Out-Null
	
	$PoolNumber = 1
	foreach($Pool in $InternalPools + $InternalServers){
		$Roles = Get-CsRoles $Pool
		$Stencil = "SFBServer"
		if($Roles -match "File Share"){$Stencil = "FileStore"}
		if($Roles -match "Front-End"){$Stencil = "FEPool"}
		if($Roles -match "Mediation"){$Stencil = "MediationServer"}
		
		$Members = @()
		$Machines = foreach($PoolMember in $Pool.Machines){
			$Members += $CsConfig.Topology.Object.Machines | where{$_.MachineId.ToString() -match $PoolMember} | Select -ExpandProperty Fqdn
		}
		$Members = $Members -join "`n`t`t"
		$Server = Add-InternalServerToPage -X1 21.5 -Y1 (17.25 - ($PoolNumber * 1.25)) -Page $CurrentPage -Shape "$Stencil" -Name $Pool.Fqdn -Roles $($Roles -join ", ") -Members $Members
		$PoolNumber++
		Add-ShapeToContainer -Shape $Server -Container "SfBLabel" -Resize 0
	}
	
	# SQL Servers label
	$SQLLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.89" -Y1 "ThePage!PageHeight * 0.625" -Name "SQLLabel" -LineColor $($Colors.Green) -Color $($Colors.White) -Height "ThePage!PageHeight * 0.375" -Width "ThePage!PageWidth * 0.17" -Transparency 30
	Set-ShapeTextFormat -Shape $SQLLabel -Size "18 pt" -Style 1 -Case 1 -HAlign 1 | Out-Null
	$SQLLabel.Text = "SQL Servers"
	
	$InstanceNumber = 1
	foreach($Instance in $SqlInstances){
		$Members = @()
		$Machines = foreach($PoolMember in $Instance.Machines){
			$Members += $CsConfig.Topology.Object.Machines | where{$_.MachineId.ToString() -match $PoolMember} | Select -ExpandProperty Fqdn
		}
		$Members = $Members -join "`n`t`t"
		Add-InternalServerToPage -X1 27.875 -Y1 (17.25 - ($InstanceNumber * 1.25)) -Page $CurrentPage -Shape "DatabaseServer" -Name $Instance.Fqdn -Roles "SQL Server" -Members $Members | Out-Null
		$InstanceNumber++		
	}
	
	# PSTN Gateways label
	$PSTNGatewaysLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.70" -Y1 "ThePage!PageHeight * 0.27" -Name "PSTNGatewaysLabel" -LineColor $($Colors.Green) -Color $($Colors.White) -Height "ThePage!PageHeight * 0.305" -Width "ThePage!PageWidth * 0.1725" -Transparency 30
	Set-ShapeTextFormat -Shape $PSTNGatewaysLabel -Size "18 pt" -Style 1 -Case 1 -HAlign 1 | Out-Null
	$PSTNGatewaysLabel.Text = "PSTN Gateways"
	
	$GatewayNumber = 1
	foreach($Gateway in $PstnGateways){
		Add-InternalServerToPage -X1 21.5 -Y1 (9 - ($GatewayNumber * 1.25)) -Page $CurrentPage -Shape "IPGateway" -Name $Gateway.Fqdn -Roles "PSTN Gateway" | Out-Null
		$GatewayNumber++		
	}
	
	# Trusted Application Server label
	$TrustedApplicationLabel = Add-ShapeToPage -Shape "Rectangle" -Page $CurrentPage -X1 "ThePage!PageWidth * 0.89" -Y1 "ThePage!PageHeight * 0.27" -Name "TrustedApplicationLabel" -LineColor $($Colors.Green) -Color $($Colors.White) -Height "ThePage!PageHeight * 0.305" -Width "ThePage!PageWidth * 0.17" -Transparency 30
	Set-ShapeTextFormat -Shape $TrustedApplicationLabel -Size "18 pt" -Style 1 -Case 1 -HAlign 1 | Out-Null
	$TrustedApplicationLabel.Text = "Trusted Application Pools and Servers"
	
	$PoolNumber = 1
	foreach($Pool in $ExternalPools + $ExternalServers){
		$Members = @()
		$Machines = foreach($PoolMember in $Pool.Machines){
			$Members += $CsConfig.Topology.Object.Machines | where{$_.MachineId.ToString() -match $PoolMember} | Select -ExpandProperty Fqdn
		}
		$Members = $Members -join "`n`t`t"
		Add-InternalServerToPage -X1 27.875 -Y1 (9 - ($PoolNumber * 1.25)) -Page $CurrentPage -Shape "ApplicationServer" -Name $Pool.Fqdn -Roles "Trusted Application" -Members $Members | Out-Null
		$PoolNumber++		
	}
	
	# Create connector arrows for associated databases and services.
	foreach($Pool in $InternalPools + $InternalServers){
		$PoolClusterSplit = [string[]]$Pool.InstalledServices[0] -split "-"
		$DBServices = $XmlTopology.Topology.Services.Service | Where {($_.DependsOn.Dependency.ServiceId.RoleName -match "Store") -and ($_.InstalledOn.ClusterId.SiteId -match $PoolClusterSplit[0]) -and($_.InstalledOn.ClusterId.Number -match $PoolClusterSplit[2])}
		$DBServiceId = $DBServices.DependsOn.Dependency.ServiceId | Where {$_.RoleName -match "Store"} | select -Unique
		$DBService = $XmlTopology.Topology.Services.Service | Where {($_.ServiceId.SiteId -eq $DBServiceId.SiteId) -and ($_.ServiceId.RoleName -eq $DBServiceId.RoleName) -and ($_.ServiceId.Instance -eq $DBServiceId.Instance)}
		$DBClusterId = $DBService.InstalledOn.SqlInstanceId.ClusterId
		$DBFqdn = $CsConfig.Topology.Object.Clusters | Where {$_.ClusterId -eq "$($DBClusterId.SiteId):$($DBClusterId.Number)"} | Select -ExpandProperty Fqdn
		
		# Create connectors from FE to SQL.
		$FESQLConnector = Add-ConnectorToPage -Page $CurrentPage -Name "$($Pool.Fqdn)SQLConnector" -BeginX 1 -BeginY 1 -EndX 2 -EndY 2  -LineWeight "0.147 pt*ThePage!PageWidth" -Color $($Colors.Red) -BeginArrow "5" -EndArrow "5" -LineType 1 -ConnectorRouteStyle 17 -ConnectorShape 2
		$FESQLConnector.Cells("BeginX").GlueTo($CurrentPage.Shapes.Item("$($Pool.Fqdn)FqdnLabel").Cells("Connections.X2"))
		$FESQLConnector.Cells("EndX").GlueTo($CurrentPage.Shapes.Item("$($DBFqdn)").Cells("Connections.X4"))

		# Create connectors from FE to Trusted Application Pools and Servers.		
		$ExternalServices = $XmlTopology.Topology.Services.Service | Where {($_.ExternalApplicationService) -and ($_.DependsOn.Dependency.ServiceId.SiteId -match $PoolClusterSplit[0]) -and($_.DependsOn.Dependency.ServiceId.Instance -match $PoolClusterSplit[2])}
		foreach($Service in $ExternalServices){
			$TrustedApplicationServer = $CsConfig.Topology.Object.Clusters | where{$_.ClusterId -match "$($Service.InstalledOn.ClusterId.SiteId):$($Service.InstalledOn.ClusterId.Number)"} | select -ExpandProperty Fqdn
			$FETAConnector = Add-ConnectorToPage -Page $CurrentPage -BeginX 1 -BeginY 1 -EndX 2 -EndY 2  -LineWeight "0.147 pt*ThePage!PageWidth" -Color $($Colors.Green) -BeginArrow "5" -EndArrow "5" -LineType 0 -ConnectorRouteStyle 17 -ConnectorShape 2
			$FETAConnector.Cells("BeginX").GlueTo($CurrentPage.Shapes.Item("$($Pool.Fqdn)FqdnLabel").Cells("Connections.X2"))
			$FETAConnector.Cells("EndX").GlueTo($CurrentPage.Shapes.Item("$($TrustedApplicationServer)").Cells("Connections.X4"))
		}
	}
	
	foreach($Trunk in $CsConfig.PolicyData.Voice.Trunk){
		# Create connectors from Mediation to PSTN Gateways.
		$PSTNFqdn = $Trunk.Identity.Replace("PstnGateway:","")
		$MedFqdn = $Trunk.MediationServer.Replace("MediationServer:","")
		$MedPSTNConnector = Add-ConnectorToPage -Page $CurrentPage -Name "$($PSTNFqdn)Connector" -BeginX 1 -BeginY 1 -EndX 2 -EndY 2  -LineWeight "0.147 pt*ThePage!PageWidth" -Color $($Colors.Orange) -BeginArrow "5" -EndArrow "5" -LineType 4 -ConnectorRouteStyle 0 -ConnectorShape 1
		$MedPSTNConnector.Cells("BeginX").GlueTo($CurrentPage.Shapes.Item("$($MedFqdn)FqdnLabel").Cells("Connections.X2"))
		$MedPSTNConnector.Cells("EndX").GlueTo($CurrentPage.Shapes.Item("$($PSTNFqdn)FqdnLabel").Cells("Connections.X2"))
	}
	
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

function Set-ShapeTextFormat {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Shape,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true)]
		[ValidateNotNullOrEmpty()]
		$Size,
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Font = "$FontIndex",
		[Parameter(Position = 3, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Color = "$($Colors.Black)",
		[Parameter(Position = 4, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Style = "0",
		[Parameter(Position = 5, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Case = "0",
		[Parameter(Position = 6, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$HAlign = "0",
		[Parameter(Position = 7, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$VAlign = "0",
		[Parameter(Position = 8, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$LeftIndent = "0 in",
		[Parameter(Position = 9, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$RightIndent = "0 in"
	)
	
	$Shape.CellsU("Char.Size").Formula = "$Size"
	$Shape.CellsU("Char.Font").Formula = "$Font"
	$Shape.CellsU("Char.Color").Formula = "THEMEGUARD($Color)"
	$Shape.CellsU("Char.Style").Formula = "$Style"
	$Shape.CellsU("Char.Case").Formula = "$Case"
	$Shape.CellsU("Para.HorzAlign").Formula = "$HAlign"
	$Shape.CellsU("VerticalAlign").Formula = "$VAlign"
	$Shape.CellsU("Para.IndLeft").Formula = "$LeftIndent"
	$Shape.CellsU("Para.IndRight").Formula = "$RightIndent"
	
	return $Shape
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

New-VisioDiagram -DataFileName $script:XmlFileName -CsConfig $CsConfig -Visible $Visible









