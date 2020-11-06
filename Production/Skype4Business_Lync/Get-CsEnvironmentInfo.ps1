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
	.\Get-CsEnvironmentInfo.ps1 [-InternalCredentials $InternalCredentials] [-EdgeCredentials $EdgeCredentials]
.INPUTS
	None. You cannot pipe objects to this script.
.PARAMETER EdgeCredentials
	A PowerShell variable containing credentials stored with Get-Credential.
#>
param([Parameter(Mandatory = $false)]
	[System.Management.Automation.PSCredential]
	[System.Management.Automation.Credential()]$InternalCredentials = [System.Management.Automation.PSCredential]::Empty,
	[System.Management.Automation.PSCredential]
	[System.Management.Automation.Credential()]$EdgeCredentials = [System.Management.Automation.PSCredential]::Empty)

$ErrorActionPreference = "ContinueSilently"

#region Variables

# Enable logging to file.
$script:EnableLogging = $true
# Set script log file.
$script:LogFileName = ".\Get-CsEnvironmentInfo.log"

#endregion Variables

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

function Create-CsEnvironmentDataFile {
	
	Update-Log "Creating base PowerShell custom object to store environment data."
	return [PSCustomObject] @{
		'TimeStamp' = [string](Get-Date)
		'FileTimeStamp' = [string](Get-Date -Format MMddyyhhmm)
		'CreatedBy' = "$env:UserDomain\$env:UserName"
		'HostName' = $env:ComputerName
		'RunAsAdmin' = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
	}
}

function Get-CsTopologyConfiguration {
	
	Update-Log "Getting current published topology information from Active Directory and the CMS."
	return [PSCustomObject] @{
		'Object' = Get-CsTopology
		'Xml' = Get-CsTopology -AsXml
	}
}

function Get-CsServerCredentials {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$InternalCredentials,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$EdgeCredentials
	)
	
	# If credentials were provided at runtime we will use those, if not we will prompt for them.
	if (!$InternalCredentials.UserName){
		Update-Log "No internal credentials provided at runtime, prompting for credentials."
		$InternalCredentials = $host.ui.PromptForCredential("Internal Credentials", "Please enter Internal server credentials.", "", "INTERNAL DOMAIN\USERNAME")
	}
	if (!$EdgeCredentials.UserName){
		Update-Log "No Edge credentials provided at runtime, prompting for credentials."
		$EdgeCredentials = $host.ui.PromptForCredential("Edge Credentials", "Please enter Edge server credentials.", "", "EDGE LOCAL ADMIN")
	}
	
	# If the user does not provide credentials at runtime and cancels the prompt it will limit the data that can be collected.
	if (!$InternalCredentials.UserName){
		$InternalCredentials = $null
		Update-Log "No internal credentials provided, internal server data collection will be limited." Warning
	}
	if (!$EdgeCredentials.UserName){
		$EdgeCredentials = $null
		Update-Log "No Edge credentials provided, Edge server data collection will be limited." Warning
	}

	return [PSCustomObject] @{
		'Internal' = $InternalCredentials
		'Edge' = $EdgeCredentials
	}
}

function Get-CsCmsConfiguration {
	
	Update-Log "Getting current CMS configuration and replication status."
	return [PSCustomObject] @{
		'ManagementConnection' = Get-CsManagementConnection
		'CMSStatus' = Get-CsManagementStoreReplicationStatus -CentralManagementStoreStatus
		'ReplicationStatus' = Get-CsManagementStoreReplicationStatus
	}
}

function Get-CsSimpleUrls {
	
	Update-Log "Getting current Simple URL configuration."
	return Get-CsSimpleUrlConfiguration
}

function Get-CsAdDomainConfiguration {
	
	Update-Log "Getting current Simple URL configuration."
	$AdDomainConfiguration = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()
	return $AdDomainConfiguration
}

function Get-CsInternalDNSRecords {
	
	Update-Log "Getting internal DNS records."
	$InternalDNSRecords = @()
	
	# Pool names
	$InternalDNSRecords += $CsConfig.Topology.Object.Clusters.fqdn
	
	# Machine Names
	$InternalDNSRecords += $CsConfig.Topology.Object.Machines.fqdn
	
	# Internal Web Services URLs
	$InternalDNSRecords += $CsConfig.Topology.Object.Clusters.InstalledServices.InternalHost
	
	# Simple URLs
	$InternalDNSRecords += (Get-CsSimpleUrlConfiguration | select -ExpandProperty SimpleUrl | Select -ExpandProperty ActiveUrl).Replace("https://","")
	
	# Predefined URLs (SIP, LyncDiscoverInternal, etc)
	$SipDomains = Get-CsSipDomain | Select -ExpandProperty Name
	foreach ($Domain in $SipDomains){
		$InternalDNSRecords += "sip.$Domain"
		$InternalDNSRecords += "lyncdiscoverinternal.$Domain"
		$InternalDNSRecords += "sipinternal.$Domain"
		$InternalDNSRecords += "_sipinternaltls._tcp.$Domain"
	}
	
	# Create array to store DNS information.
	$InternalDNSConfiguration = @{}
	
	# Enumerate DNS records, resolve them, and then return the results.
	foreach($DNSRecord in $InternalDNSRecords | Select -Unique){
		$DNSRecordDetails = Resolve-DNSRecord $DNSRecord
		$InternalDNSConfiguration.Add("$DNSRecord","$DNSRecordDetails")
	}
	
	return $InternalDNSConfiguration
}

function Get-CsExternalDNSRecords {
	
	Update-Log "Getting external DNS records."
	$ExternalDNSRecords = @()
	
	# External Web Services URLs
	$ExternalDNSRecords += $CsConfig.Topology.Object.Clusters.InstalledServices.ExternalHost
	
	# Edge server external FQDNs
	$ExternalDNSRecords += $CsConfig.Topology.Object.Clusters.InstalledServices.NetPorts | ?{($_.Service -match "Edge") -and ($_.NetInterfaceId -match "External")} | select -Unique -ExpandProperty ConfiguredFqdn
	
	# Simple URLs
	$ExternalDNSRecords += (Get-CsSimpleUrlConfiguration | select -ExpandProperty SimpleUrl | Select -ExpandProperty ActiveUrl).Replace("https://","")
	
	# Predefined URLs (SIP, LyncDiscoverInternal, etc)
	$SipDomains = Get-CsSipDomain | Select -ExpandProperty Name
	foreach ($Domain in $SipDomains){
		$ExternalDNSRecords += "sip.$Domain"
		$ExternalDNSRecords += "lyncdiscover.$Domain"
		$ExternalDNSRecords += "sipexternal.$Domain"
		$ExternalDNSRecords += "_sip._tls.$Domain"
		$ExternalDNSRecords += "_sipfederationtls._tcp.$Domain"
	}
	
	# Create array to store DNS information.
	$ExternalDNSConfiguration = @{}
	
	# Enumerate DNS records, resolve them, and then return the results.
	foreach($DNSRecord in $ExternalDNSRecords | Select -Unique){
		$DNSRecordDetails = Resolve-DNSRecord $DNSRecord -DNSView External
		$ExternalDNSConfiguration.Add("$DNSRecord","$DNSRecordDetails")
	}
	
	return $ExternalDNSConfiguration
}

function Resolve-DNSRecord {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true, HelpMessage = "No DNS record provided.")]
		[ValidateNotNullOrEmpty()]
		[string] $DNSRecord,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[ValidateSet("Internal", "External")]
		[string] $DNSView = "Internal"
	)
	
	# If the DNS record contains an underscore flag it as a SRV Record.
	$SRVRecord = $DNSRecord -match "_"
	
	try {
		
		# Reset the NSLookup and NSLookupResults command variables.
		$NSLookup = $null
		$NSLookupResults = $null
		
		# Set the base component of the NSLookup command.
		$NSLookup = "nslookup.exe"
		
		# If this is a SRV record then append the record type setting.
		if($SRVRecord){$NSLookup = "$NSLookup -timeout=1 -type=srv"}
		
		# Append the DNS Record to be resolved.
		$NSLookup = "$NSLookup $DNSRecord"
		
		# If the DNS view is external then append the public DNS server address to use for name resolution.
		if($DNSView -eq "External"){$NSLookup = "$NSLookup 8.8.8.8"}
		
		# Invoke the expression and grab the results.
		$NSLookupResults = Invoke-Expression "$NSLookup"
		
		if ($SRVRecord){
			# Reset DNSRecordTarget variable as an empty array.
			$DNSRecordTarget = @()
			# Enumerate stdout results from the NSLookup command.
			for ($i = 4; $i -lt $NSLookupResults.Count; $i++){$DNSRecordTarget += ($NSLookupResults[$i].Replace("  ","")).Trim() + "`n"}
		} else {
			# Reset DNSRecordTarget variable as an empty array.
			$DNSRecordTarget = @()
			# Enumerate stdout results from the NSLookup command.
			for ($i = 4; $i -lt ($NSLookupResults.Count - 1); $i++){
				$Address = $NSLookupResults[$i].Replace("Addresses:","")
				$Address = $Address.Replace("Address:","")
				$DNSRecordTarget += $Address.Trim()
			}
		}
	}
	catch {
		Update-Log $Error[0].Exception.Message WARNING
		$DNSRecordTarget = "ERROR"
	}
	
	# Return the lookup results.
	return $DNSRecordTarget
}

function Get-CsPolicyConfiguration {
	
	# Create PowerShell custom object to hold policy configuration data.
	$PolicyData = New-Object PSCustomObject
	Update-Log "Gathering Edge, Mobility, and Federation configuration data..."
	$EdgeCmdlets = @{
		"AccessEdgeConfiguration" = "Get-CsAccessEdgeConfiguration"
		"AllowedDomain" = "Get-CsAllowedDomain"
		"BlockedDomain" = "Get-CsBlockedDomain"
		"ExternalAccessPolicy" = "Get-CsExternalAccessPolicy"
		"MobilityPolicy" = "Get-CsMobilityPolicy"
		"McxConfiguration" = "Get-CsMcxConfiguration"
	}
	$EdgeConfig = New-Object PSObject
	foreach ($PSHCmdlet in $EdgeCmdlets.GetEnumerator() | Sort-Object Key) {Add-Member -InputObject $EdgeConfig -MemberType NoteProperty -Name ($PSHCmdlet.Key) -Value (Invoke-Expression ($PSHCmdlet.Value) -EA SilentlyContinue)}
	Add-Member -InputObject $PolicyData -MemberType NoteProperty -Name ExternalConfig -Value $EdgeConfig

	Update-Log "Gathering Archiving and Monitoring configuration data..."
	$ArchMonCmdlets = @{
		"ArchivingConfiguration" = "Get-CsArchivingConfiguration"
		"ArchivingPolicy" = "Get-CsArchivingPolicy"
		"QoEConfiguration" = "Get-CsQoEConfiguration"
	}
	$ArchMonConfig = New-Object PSObject
	foreach ($PSHCmdlet in $ArchMonCmdlets.GetEnumerator() | Sort-Object Key) {Add-Member -InputObject $ArchMonConfig -MemberType NoteProperty -Name ($PSHCmdlet.Key) -Value (Invoke-Expression ($PSHCmdlet.Value) -EA SilentlyContinue)}
	Add-Member -InputObject $PolicyData -MemberType NoteProperty -Name ArchivingConfig -Value $ArchMonConfig

	Update-Log "Gathering Call Admission Control configuration data..."
	$CACCmdlets = @{
		"BandwidthPolicyServiceConfiguration" = "Get-CsBandwidthPolicyServiceConfiguration"
		"NetworkBandwidthPolicyProfile" = "Get-CsNetworkBandwidthPolicyProfile"
		"NetworkConfiguration" = "Get-CsNetworkConfiguration"
		"NetworkInterRegionRoute" = "Get-CsNetworkInterRegionRoute"
		"NetworkInterSitePolicy" = "Get-CsNetworkInterSitePolicy"
		"NetworkRegion" = "Get-CsNetworkRegion"
		"NetworkRegionLink" = "Get-CsNetworkRegionLink"
		"NetworkSite" = "Get-CsNetworkSite"
		"NetworkSubnet" = "Get-CsNetworkSubnet"
	}
	$CACConfig = New-Object PSObject
	foreach ($PSHCmdlet in $CACCmdlets.GetEnumerator() | Sort-Object Key) {Add-Member -InputObject $CACConfig -MemberType NoteProperty -Name ($PSHCmdlet.Key) -Value (Invoke-Expression ($PSHCmdlet.Value) -EA SilentlyContinue)}
	Add-Member -InputObject $PolicyData -MemberType NoteProperty -Name CAC -Value $CACConfig

	Update-Log "Gathering Location Information Service configuration data..."
	$LISCmdlets = @{
		"LisCivicAddress" = "Get-CsLisCivicAddress"
		"LisLocation" = "Get-CsLisLocation"
		"LisPort" = "Get-CsLisPort"
		"LisServiceProvider" = "Get-CsLisServiceProvider"
		"LisSubnet" = "Get-CsLisSubnet"
		"LisSwitch" = "Get-CsLisSwitch"
		"LisWirelessAccessPoint" = "Get-CsLisWirelessAccessPoint"
	}
	$LISConfig = New-Object PSObject
	foreach ($PSHCmdlet in $LISCmdlets.GetEnumerator() | Sort-Object Key) {Add-Member -InputObject $LISConfig -MemberType NoteProperty -Name ($PSHCmdlet.Key) -Value (Invoke-Expression ($PSHCmdlet.Value) -EA SilentlyContinue)}
	Add-Member -InputObject $PolicyData -MemberType NoteProperty -Name LIS -Value $LISConfig

	Update-Log "Gathering Conferencing configuration data..."
	$ConferencingCmdlets = @{
		"ConferenceDisclaimer" = "Get-CsConferenceDisclaimer"
		"ConferencingConfiguration" = "Get-CsConferencingConfiguration"
		"ConferencingPolicy" = "Get-CsConferencingPolicy"
		"MeetingConfiguration" = "Get-CsMeetingConfiguration"
		"DialinConferencingAccessNumber" = "Get-CsDialinConferencingAccessNumber"
		"DialinConferencingConfiguration" = "Get-CsDialinConferencingConfiguration"
		"DialinConferencingLanguageList" = "Get-CsDialinConferencingLanguageList"
	}
	$ConferencingConfig = New-Object PSObject
	foreach ($PSHCmdlet in $ConferencingCmdlets.GetEnumerator() | Sort-Object Key) {Add-Member -InputObject $ConferencingConfig -MemberType NoteProperty -Name ($PSHCmdlet.Key) -Value (Invoke-Expression ($PSHCmdlet.Value) -EA SilentlyContinue)}
	Add-Member -InputObject $PolicyData -MemberType NoteProperty -Name Conferencing -Value $ConferencingConfig

	Update-Log "Gathering Enterprise Voice configuration data..."
	$VoiceCmdlets = @{
		"PstnUsage" = "Get-CsPstnUsage"
		"RoutingConfiguration" = "Get-CsRoutingConfiguration"
		"Trunk" = "Get-CsTrunk"
		"TrunkConfiguration" = "Get-CsTrunkConfiguration"
		"VoiceConfiguration" = "Get-CsVoiceConfiguration"
		"VoiceNormalizationRule" = "Get-CsVoiceNormalizationRule"
		"VoiceRoute" = "Get-CsVoiceRoute"
		"VoicePolicy" = "Get-CsVoicePolicy"
		"PinPolicy" = "Get-CsPinPolicy"
		"CpsConfiguration" = "Get-CsCpsConfiguration"
		"MediaConfiguration" = "Get-CsMediaConfiguration"
		"DialPlan" = "Get-CsDialPlan"
	}
	$VoiceConfig = New-Object PSObject
	foreach ($PSHCmdlet in $VoiceCmdlets.GetEnumerator() | Sort-Object Key) {Add-Member -InputObject $VoiceConfig -MemberType NoteProperty -Name ($PSHCmdlet.Key) -Value (Invoke-Expression ($PSHCmdlet.Value) -EA SilentlyContinue)}
	Add-Member -InputObject $PolicyData -MemberType NoteProperty -Name Voice -Value $VoiceConfig

	Update-Log "Gathering Response Group configuration data..."
	$RGSCmdlets = @{
		"RgsAgentGroup" = "Get-CsRgsAgentGroup"
		"RgsConfiguration" = "Get-CsService -ApplicationServer | ForEach-Object {Get-CsRgsConfiguration -Identity `$_.Identity}"
		"RgsHolidaySet" = "Get-CsRgsHolidaySet"
		"RgsHoursOfBusiness" = "Get-CsRgsHoursOfBusiness"
		"RgsQueue" = "Get-CsRgsQueue"
		"RgsWorkflow" = "Get-CsRgsWorkflow"
	}
	$RGSConfig = New-Object PSObject
	foreach ($PSHCmdlet in $RGSCmdlets.GetEnumerator() | Sort-Object Key) {Add-Member -InputObject $RGSConfig -MemberType NoteProperty -Name ($PSHCmdlet.Key) -Value (Invoke-Expression ($PSHCmdlet.Value) -EA SilentlyContinue)}
	Add-Member -InputObject $PolicyData -MemberType NoteProperty -Name RGS -Value $RGSConfig

	Update-Log "Gathering Lync Policy configuration data..."
	$PolicyCmdlets = @{
		"AddressBookConfiguration" = "Get-CsAddressBookConfiguration"
		"ClientPolicy" = "Get-CsClientPolicy"
		"ClientVersionPolicy" = "Get-CsClientVersionPolicy"
		"FileTransferFilterConfiguration" = "Get-CsFileTransferFilterConfiguration"
		"IMFilterConfiguration" = "Get-CsImFilterConfiguration"
		"PresencePolicy" = "Get-CsPresencePolicy"
		"PrivacyConfiguration" = "Get-CsPrivacyConfiguration"
		"HealthMonitoringConfiguration" = "Get-CsHealthMonitoringConfiguration"
	}
	$PolicyConfig = New-Object PSObject
	foreach ($PSHCmdlet in $PolicyCmdlets.GetEnumerator() | Sort-Object Key) {Add-Member -InputObject $PolicyConfig -MemberType NoteProperty -Name ($PSHCmdlet.Key) -Value (Invoke-Expression ($PSHCmdlet.Value) -EA SilentlyContinue)}
	Add-Member -InputObject $PolicyData -MemberType NoteProperty -Name Policy -Value $PolicyConfig
	
	# Return the policy data object.
	return $PolicyData
}

function Get-CsServerData {
	
		
	Update-Log "Getting list of servers from topology for additional information collection. This might take a while."
	# Get list of all internal servers that are not PSTN Gateways.
	$InternalServerList = $CsConfig.EnvironmentData.PoolData | Where {($_.Services -notmatch "PstnGateway") -and ($_.Services -notmatch "EdgeServer") -and ($_.Services -notmatch "TrustedApplicationPool")} | Select -ExpandProperty Computers
	# Get list of all Edge servers that are in the topology because the connection method and data collection is different.
	$EdgeServerList = $CsConfig.EnvironmentData.PoolData | Where {$_.Services -match "EdgeServer"} | Select -ExpandProperty Computers
	# Get list of all Trusted Application servers that are in the topology because the data collected is less detailed.
	$TrustedApplicationServerList = $CsConfig.EnvironmentData.PoolData | Where {$_.Services -match "TrustedApplicationPool"} | Select -ExpandProperty Computers
	
	$ServerDetailsHashTable = @{}
	
	# Enumerate internal servers and connect to them to gather detailed information.
	foreach ($Server in $InternalServerList){
		$ServerDetails = Get-CsServerDetails $Server -Credentials $script:Credentials.Internal
		$ServerDetailsHashTable.Add($Server,$ServerDetails)
	}
	
	# Enumerate Edge servers and connect to them to gather detailed information.
	foreach ($Server in $EdgeServerList){
		$ServerDetails = Get-CsServerDetails $Server -ServerType Edge -Credentials $script:Credentials.Edge
		$ServerDetailsHashTable.Add($Server,$ServerDetails)
	}
	
	# Enumerate Trusted Application servers and connect to them to gather detailed information.
	foreach ($Server in $TrustedApplicationServerList){
		$ServerDetails = Get-CsServerDetails $Server -Credentials $script:Credentials.Internal
		$ServerDetailsHashTable.Add($Server,$ServerDetails)
	}
	
	return $ServerDetailsHashTable
}

function Get-CsServerDetails {
	[CmdletBinding(SupportsShouldProcess = $True, SupportsPaging = $True)]
	param(
		[Parameter(Position = 0, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $true, HelpMessage = "No server fqdn or name provided.")]
		[ValidateNotNullOrEmpty()]
		[string] $Fqdn,
		[Parameter(Position = 1, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		[ValidateSet("Internal", "Edge", "TrustedApplication")]
		[string] $ServerType = "Internal",
		[Parameter(Position = 2, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, Mandatory = $false)]
		$Credentials
	)
	
	Update-Log "Collecting additional detailed information on $Fqdn."

	# Create custom object to store server details.
	$ServerDetails = New-Object PSCustomObject
	
	# Lookup server in topology for additional details.
	$ServerTopologyDetails = $CsConfig.Topology.Object.Machines | ? {$_.Fqdn -match "$Fqdn"}
	
	# Get CPU details.
	try{
		$CmdToRun = "Get-WmiObject Win32_Processor -ComputerName $Fqdn"
		if(($ServerType -eq "Edge") -and ($Credentials)){$CmdToRun = "$CmdToRun -Credential `$Credentials"}
		$CPUInfo = Invoke-Expression "$CmdToRun"
	}
	catch{
		$CPUInfo = "ERROR"
		Update-Log "Unable to query remote machine $Fqdn for CPU details." WARNING
	}
	
	# Get OS configuration and details.
	try{
		$CmdToRun = "Get-WmiObject Win32_OperatingSystem -ComputerName $Fqdn"
		if(($ServerType -eq "Edge") -and ($Credentials)){$CmdToRun = "$CmdToRun -Credential `$Credentials"}
		$OSInfo = Invoke-Expression "$CmdToRun"
	}
	catch{
		$OSInfo = "ERROR"
		Update-Log "Unable to query remote machine $Fqdn for OS details." WARNING
	}
	
	# Get RAM configuration and details.
	try{
		$CmdToRun = "Get-WmiObject CIM_PhysicalMemory -ComputerName $Fqdn"
		if(($ServerType -eq "Edge") -and ($Credentials)){$CmdToRun = "$CmdToRun -Credential `$Credentials"}
		$RAMInfo = Invoke-Expression "$CmdToRun"
	}
	catch{
		$RAMInfo = "ERROR"
		Update-Log "Unable to query remote machine $Fqdn for RAM details." WARNING
	}
	
	# Get Software Version details.
	try{
		$CmdToRun = "Get-WmiObject -query 'Select * from Win32_Product' -ComputerName $Fqdn"
		if(($ServerType -eq "Edge") -and ($Credentials)){$CmdToRun = "$CmdToRun -Credential `$Credentials"}
		$SWVersionInfo = Invoke-Expression "$CmdToRun"
	}
	catch{
		$SWVersionInfo = "ERROR"
		Update-Log "Unable to query remote machine $Fqdn for Software Version details." WARNING
	}
	
	# Get Hotfix configuration and details.
	try{
		$CmdToRun = "Get-Hotfix -ComputerName $Fqdn"
		if(($ServerType -eq "Edge") -and ($Credentials)){$CmdToRun = "$CmdToRun -Credential `$Credentials"}
		$HotfixInfo = Invoke-Expression "$CmdToRun"
	}
	catch{
		$HotfixInfo = "ERROR"
		Update-Log "Unable to query remote machine $Fqdn for Hotfix details." WARNING
	}
	
	# Get CS Service details.
	try{
		$CmdToRun = "Invoke-Command -ComputerName $Fqdn -ScriptBlock {Get-CsWindowsService} -Credential `$Credentials"
		$ServiceInfo = Invoke-Expression "$CmdToRun"
	}
	catch{
		$ServiceInfo = "ERROR"
		Update-Log "Unable to query remote machine $Fqdn for CS Service details." WARNING
	}
	
	# Using Get-CsCertificate with Invoke-Command against a remote system returns an empty AlternativeNames property.
	# .NET must be used to read the X.509 configuration in order to get the proper information.
	try{
		# If this is an Edge server we need to map the C$ admininistrative share to provide implicit credentials for certificate store lookups instead of using the current user credentials.
		if(($ServerType -eq "Edge") -and ($Credentials)){
			try{
				net use \\$Fqdn\c`$ "$($Credentials.GetNetworkCredential().Password)" /user:"$($Credentials.GetNetworkCredential().Username)" | Out-Null
			}
			catch{
				Update-Log "Unable to map drive to $Fqdn for certificate lookup." WARNING
			}
		}
		$Certificates = @{}
		$CertUse = $null
		$CertReadOnly = [System.Security.Cryptography.X509Certificates.OpenFlags]"ReadOnly"
		$CertLmStore = [System.Security.Cryptography.X509Certificates.StoreLocation]"LocalMachine"
		$CertStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("\\$Fqdn\my",$CertLmStore)
		$CertStore.Open($CertReadOnly)
		
		$CertResults = $Null
		$CertResults = $CertStore.Certificates
		
		$CmdToRun = "Invoke-Command -ComputerName $Fqdn -ScriptBlock {Get-CsCertificate} -Credential `$Credentials"
		$CertificateInfo = Invoke-Expression "$CmdToRun"
		
		foreach ($Certificate in $CertificateInfo){
			$CertAltNames = $null
			$CertSanExtension = $null
			$CurrentCert = $null
			
			$CurrentCert = $CertResults | ? {$_.Thumbprint -match $Certificate.Thumbprint}
			$CertSanExtension = $CurrentCert.Extensions | Where {$_.Oid.FriendlyName -match "subject alternative name"}
			if ($CertSanExtension){
				$CertAltNames = $CertSanExtension.Format(1)
				$CertAltNames = $CertAltNames.Replace("DNS Name=", "")
				$tmpCertAltNames = $CertAltNames.Replace("`r`n", " ")
				$CertAltNames = $tmpCertAltNames.Split(" ")
			}
			
			Add-Member -InputObject $Certificate -MemberType NoteProperty -Name AlternativeNames -Value $CertAltNames -Force
		}
		
		# Remove the mapped drive to the Edge server.
		try{
			if (($ServerType -eq "Edge") -and ($Credentials)){net use /Delete \\$Fqdn\c`$ | Out-Null}
		}
		catch{
		}

	}
	catch{
		$CertificateInfo = "ERROR"
		Update-Log "Unable to query remote machine $Fqdn for Certificate details." WARNING
	}
	
	Add-Member -InputObject $ServerDetails -MemberType NoteProperty -Name FQDN -Value $Fqdn
	Add-Member -InputObject $ServerDetails -MemberType NoteProperty -Name CPUInfo -Value $CPUInfo
	Add-Member -InputObject $ServerDetails -MemberType NoteProperty -Name OSInfo -Value $OSInfo
	Add-Member -InputObject $ServerDetails -MemberType NoteProperty -Name RAMInfo -Value $RAMInfo
	Add-Member -InputObject $ServerDetails -MemberType NoteProperty -Name SWVersionInfo -Value $SWVersionInfo
	Add-Member -InputObject $ServerDetails -MemberType NoteProperty -Name HotfixInfo -Value $HotfixInfo
	Add-Member -InputObject $ServerDetails -MemberType NoteProperty -Name ServiceInfo -Value $ServiceInfo
	Add-Member -InputObject $ServerDetails -MemberType NoteProperty -Name CertificateInfo -Value $CertificateInfo
	Add-Member -InputObject $ServerDetails -MemberType NoteProperty -Name TopologyDetails -Value $ServerTopologyDetails
	
	return $ServerDetails
}

function Get-CsDBMirrorData {
		
	Update-Log "Getting database mirroring details and cofiguration."
	# Create a hash table to store mirror details.
	$DBMirrorData = @{}
	
	# Get all pools that are not Edge, PSTN Gateways, or External Trusted Application servers.
	$PoolList = $CsConfig.EnvironmentData.PoolData | Where {($_.Services -notmatch "PstnGateway") -and ($_.Services -notmatch "EdgeServer") -and ($_.Services -notmatch "TrustedApplicationPool")} | Select -ExpandProperty Fqdn

	foreach($Pool in $PoolList){
		$MirrorDetails = Get-CsDatabaseMirrorState -PoolFqdn $Pool
		$DBMirrorData.Add($Pool,$MirrorDetails)
	}
	
	return $DBMirrorData
}

function Export-CsEnvironmentData {

	Update-Log "Exporting data to XML..."
	# Get current path.
	[string]$CurrentPath = Get-Location
	
	# Contruct filename for zip package.
	[string]$ZipFilename = "$CurrentPath\$($CsConfig.EnvironmentData.AdDomain.Name) CS_Env_Data-$($CsConfig.FileTimeStamp).zip"
	Set-Content $ZipFilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))
		(dir $ZipFilename).IsReadOnly = $false
	$ShellApp = New-Object -COM Shell.Application
	$ZipPackage = $ShellApp.NameSpace($ZipFilename)
	$CsConfig | Export-Clixml -Path "$CurrentPath\$($CsConfig.EnvironmentData.AdDomain.Name) CS_Env_Data-$($CsConfig.FileTimeStamp).xml"
	$ZipPackage.MoveHere("$CurrentPath\$($CsConfig.EnvironmentData.AdDomain.Name) CS_Env_Data-$($CsConfig.FileTimeStamp).xml")

	Update-Log "Finished gathering CS configuration and policy data. All information is stored in ""$ZipFilename""."
}

# Prompt for internal and Edge server credentials.
$script:Credentials = Get-CsServerCredentials -InternalCredentials $InternalCredentials -EdgeCredentials $EdgeCredentials

# Create PowerShell custom object for configuration information.
$script:CsConfig = Create-CsEnvironmentDataFile

# Grab topology as Object and as XML Object and attach them to CsConfig object.
$TopologyConfiguration = Get-CsTopologyConfiguration
Add-Member -InputObject $CsConfig -MemberType NoteProperty -Name Topology -Value $TopologyConfiguration

# Create place holder for environment configuration and attach it to CsConfig object.
$EnvironmentData = New-Object PSCustomObject
Add-Member -InputObject $CsConfig -MemberType NoteProperty -Name EnvironmentData -Value $EnvironmentData

# Get list of RTC users, but do not return or store any personal information.
$UserData = Get-CsUser | Select Enabled,EnterPriseVoiceEnabled,RegistrarPool
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name UserData -Value $UserData

# Get SIP Domains in environment.
$SipDomains = Get-CsSipDomain
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name SipDomains -Value $SipDomains

# Get current CMS configuration and status.
$CmsConfiguration = Get-CsCmsConfiguration
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name CMS -Value $CmsConfiguration

# Get current Simple URL configuration.
$SimpleUrlConfiguration = Get-CsSimpleUrls
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name SimpleUrls -Value $SimpleUrlConfiguration

# Get SIP Domains in environment.
$AdDomainConfiguration = Get-CsAdDomainConfiguration
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name AdDomain -Value $AdDomainConfiguration

# Get Internal DNS Records in environment.
$InternalDNS = Get-CsInternalDNSRecords
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name InternalDNS -Value $InternalDNS

# Get External DNS Records in environment.
$ExternalDNS = Get-CsExternalDNSRecords
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name ExternalDNS -Value $ExternalDNS

# Get CS Service details and configuration.
$ServiceData = Get-CsService
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name ServiceData -Value $ServiceData

# Get pool details and configuration.
$PoolData = Get-CsPool
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name PoolData -Value $PoolData

# Get pool details and configuration.
$ServerData = Get-CsServerData
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name ServerData -Value $ServerData

# Get Database Mirroring details and configuration.
$DBMirrorData = Get-CsDBMirrorData
Add-Member -InputObject $CsConfig.EnvironmentData -MemberType NoteProperty -Name DBMirrorData -Value $DBMirrorData

# Get RTC policies configuration settings in the current environment.
$PolicyConfig = Get-CsPolicyConfiguration
Add-Member -InputObject $CsConfig -MemberType NoteProperty -Name PolicyData -Value $PolicyConfig

# Export data to XML and store it in a zip file.
Export-CsEnvironmentData




