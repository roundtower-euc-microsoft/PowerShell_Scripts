# Generate HyperV VM Health Report
# Version - 3.2
# Created By - Shabarianth Ramadasan - InsideVirtualization.Com
# shabarinath@insidevirtualization.com
# Date - 6/24/2014
Import-Module -Name FailoverClusters, hyper-v
$ResultFile = "HyperV-VM-Report.htm" 
New-Item -ItemType File $ResultFile -Force
[Array]$WarningLevel = "#77FF5B","#F4FA58","#FF8000","#FF0040"
Function fWriteHtmlHeader 
	{ 
	param($FileName) 
	$date = ( get-date ).ToString('yyyy/MM/dd') 
	Add-Content $FileName "<html>" 
	Add-Content $FileName "<head>" 
	Add-Content $FileName "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>" 
	Add-Content $FileName '<title>Hyper-V Dashboard</title>' 
	Add-Content $FileName '<STYLE TYPE="text/css">' 
	Add-Content $FileName  "<!--" 
	Add-Content $FileName  "td {" 
	Add-Content $FileName  "font-family: Tahoma;" 
	Add-Content $FileName  "font-size: 11px;" 
	Add-Content $FileName  "border-top: 2px solid #999999;" 
	Add-Content $FileName  "border-right: 2px solid #999999;" 
	Add-Content $FileName  "border-bottom: 2px solid #999999;" 
	Add-Content $FileName  "border-left: 2px solid #999999;" 
	Add-Content $FileName  "}" 
	Add-Content $FileName  "body {" 
    	Add-Content $FileName  "margin-left: 5px;" 
	Add-Content $FileName  "margin-top: 5px;" 
	Add-Content $FileName  "margin-right: 5px;" 
	Add-Content $FileName  "margin-bottom: 5px;" 
	Add-Content $FileName  "" 
	Add-Content $FileName  "table {" 
	Add-Content $FileName  "border: thin solid #000000;" 
	Add-Content $FileName  "}" 
	Add-Content $FileName  "-->" 
	Add-Content $FileName  "</style>" 
	Add-Content $FileName "</head>" 
	Add-Content $FileName "<body>" 
	Add-Content $FileName  "<table width='100%'>" 
	Add-Content $FileName  "<tr bgcolor='#2F0B3A'>" 
	Add-Content $FileName  "<td colspan='30' height='20' align='center'>" 
	Add-Content $FileName  "<font face='tahoma' color='#FFFF00' size='4'><strong>Hyper-V - VM Dashboard -  $date</strong></font>" 
	Add-Content $FileName  "</td>" 
	Add-Content $FileName  "</tr>" 
	Add-Content $FileName  "</table>" 
    } 
Function fWriteVMTableHeader 
	{ 
	Param($FileName) 
	Add-Content $FileName  "<table width='100%'>"
 	Add-Content $FileName "<tr bgcolor=#BE81F7>" 
	Add-Content $FileName "<td width='10%' align=center>VM</td>"
	Add-Content $FileName "<td width='5%' align=center>Up-Time</td>"
	Add-Content $FileName "<td width='5%' align=center>IC Version</td>"
	Add-Content $FileName "<td width='4%' align=center>Clustered</td>"
	Add-Content $FileName "<td width='4%' align=center>vProcessor</td>"
	Add-Content $FileName "<td width='6%' align=center>vRAM-StartUp</td>"
    	Add-Content $FileName "<td width='6%' align=center>vRAM-Min</td>"
    	Add-Content $FileName "<td width='6%' align=center>vRAM-Max</td>"
    	Add-Content $FileName "<td width='6%' align=center>vRAM-Avg</td>"
	Add-Content $FileName "<td width='6%' align=center>vDisk1-Storage</td>"
	Add-Content $FileName "<td width='6%' align=center>vDisk1-Allocated</td>"
	Add-Content $FileName "<td width='6%' align=center>vDisk1-Usage</td>"
	Add-Content $FileName "<td width='6%' align=center>vDisk2-Storage</td>"
	Add-Content $FileName "<td width='6%' align=center>vDisk2-Allocated</td>"
	Add-Content $FileName "<td width='6%' align=center>vDisk2-Usage</td>"
	Add-Content $FileName "<td width='6%' align=center>vNic</td>"
	Add-Content $FileName "<td width='6%' align=center>FirstSnapShotDate</td>"
	Add-Content $FileName "</tr>" 
	}
Function fWriteHostStatusTableHeader 
	{
	Param($FileName) 
	Add-Content $FileName  "<table width='100%'>"
 	Add-Content $FileName "<tr bgcolor=#BE81F7>" 
	Add-Content $FileName "<td width='10%' align=center>Server Name</td>"
	Add-Content $FileName "<td width='10%' align=center>Total VMs</td>"
	Add-Content $FileName "<td width='10%' align=center>Total Physical Memory</td>"
	Add-Content $FileName "<td width='10%' align=center>Available Physical Memory</td>"
	Add-Content $FileName "<td width='10%' align=center>Total VM Startup Memory</td>"
    	Add-Content $FileName "<td width='10%' align=center>Toatl VM Max Memory</td>"
    	Add-Content $FileName "<td width='10%' align=center>Physical Processors</td>"
	Add-Content $FileName "<td width='10%' align=center>Cores</td>"
	Add-Content $FileName "<td width='10%' align=center>Logical Processors</td>"
	Add-Content $FileName "<td width='10%' align=center>Total vProcs</td>"
	Add-Content $FileName "</tr>" 
	}
Function fWriteStorageTableHeader 
	{
	Param($FileName) 
	Add-Content $FileName  "<table width='100%'>"
 	Add-Content $FileName "<tr bgcolor=#BE81F7>" 
	Add-Content $FileName "<td width='10%' align=center>Volume Name</td>" 
	Add-Content $FileName "<td width='10%' align=center>Total Capacity</td>"
	Add-Content $FileName "<td width='10%' align=center>Current Free Space</td>"
	Add-Content $FileName "<td width='10%' align=center>Free Space (%)</td>"
    	Add-Content $FileName "<td width='10%' align=center>VHDX-Size-Allocated</td>"
    	Add-Content $FileName "<td width='10%' align=center>VHDX-Size-Actual</td>"
	Add-Content $FileName "<td width='10%' align=center>Storage-Over-Provisioned</td>"
	Add-Content $FileName "<td width='10%' align=center>Storage-Over-Provisioned (%)</td>"
	Add-Content $FileName "</tr>" 
	}
Function fWriteHtmlFooter 
	{ 
	Param($FileName)  
	Add-Content $FileName "</body>" 
	Add-Content $FileName "</html>" 
	} 
Function fWriteSubRowNodeName
	{
	Param ($FileName, $nodeName, $TotMem, $AvailMem, $AvailMemPC, $hostMemHealth)
    	Add-Content $FileName  "<table width='100%'>" 
	Add-Content $FileName  "<tr height='20' bgcolor='#000000'>" 
	Add-Content $FileName  "<td width = '64%' size='3' align=center><strong>$nodeName</strong></td>" 
	Add-Content $FileName "<td width='8%' align=center><font color='White'>Total Memory</Font></td>"
	Add-Content $FileName "<td width='4%' align=center><font color='White'>$TotMem GB</Font></td>"
	Add-Content $FileName "<td width='8%' align=center><font color='White'>Available Memory</Font></td>"
	Add-Content $FileName "<td width='4%' align=center><font color='White'>$AvailMem GB</Font></td>"
	Add-Content $FileName "<td width='8%' align=center><font color='White'>Available Memory (%)</Font></td>"
	Add-Content $FileName "<td width='4%' align=center><strong><font size='4' color='$hostMemHealth'>$AvailMemPC % </font></strong></td>"
	Add-Content $FileName  "</tr>"
	Add-Content $FileName  "</table>"
    	}
Function fWriteSubHeadingCSVDetails
	{
	Param ($FileName, $cluname)
    	Add-Content $FileName  "<table width='100%'>" 
	Add-Content $FileName  "<tr colspan='1' height='20' align='center' bgcolor='#282828'>" 
	Add-Content $FileName  "<td width = '100%' color='#0101DF' size='2' align=center><strong>Storage Report - $CluName</strong></td>" 
	Add-Content $FileName  "</tr>" 
	Add-Content $FileName  "</table>" 
    	}
Function fWriteSubHeadingHostDetails
	{
	Param ($FileName, $cluname)
    	Add-Content $FileName  "<table width='100%'>" 
	Add-Content $FileName  "<tr colspan='1' height='20' align='center' bgcolor='#282828'>" 
	Add-Content $FileName  "<td width = '100%' color='#0101DF' size='2' align=center><strong>Host Report - $CluName</strong></td>" 
	Add-Content $FileName  "</tr>" 
	Add-Content $FileName  "</table>" 
    	}
Function fWriteVMInfo
	{ 
	Param($FileName, $vmname, $utime, $ic, $clusterrole, $vProc, $Startmem, $MinMem, $MaxMem, $AvgMem, $hostmemhealth, $vd1storage, $vd1, $vdu1, $vd1StorageHealth, $vdtype1, $vd2Storage, $vd2, $vdu2, $vd2StorageHealth, $vdtype2, $vNetworkInterfaceType, $SSDate, $ICStatus, $vhealth)
	Add-Content $FileName "<tr bgcolor=#77FF5B>"
	Add-Content $FileName "<td width='10%' align=center>$vmname</td>" 
	Add-Content $FileName "<td width='5%' align=center>$utime</td>"
	If ($ICStatus -match "Update required")
		{
		Add-Content $FileName "<td width='5%'  BGCOLOR='#F4FA58' align=center><strong>$ic</strong></td>"
		}
	Else
		{
		Add-Content $FileName "<td width='5%' align=center>$ic</td>"
		}
	Add-Content $FileName "<td width='4%' align=center>$clusterrole</td>" 
	Add-Content $FileName "<td width='4%' align=center>$vProc</td>"
	Add-Content $FileName "<td width='5%' align=center>$StartMem GB</td>"
	If ($MinMem -eq "DM Disabled")
		{
		Add-Content $FileName "<td width='6%' align=center>$MinMem</td>"
    		Add-Content $FileName "<td width='6%' align=center>$MaxMem</td>"
		}
	Else
		{
    		Add-Content $FileName "<td width='6%' align=center>$MinMem GB</td>"
    		Add-Content $FileName "<td width='6%' align=center>$MaxMem GB</td>"
		}
	If ($AvgMem -eq "NA")
		{
    		Add-Content $FileName "<td width='6%' align=center>$AvgMem</td>"
		}
	Else
		{
    		Add-Content $FileName "<td width='6%' align=center>$AvgMem GB</td>"
		}
	Add-Content $FileName "<td width='6%' BGCOLOR='$vd1StorageHealth' align=center>$vd1Storage</td>"
	If (($vdtype1 -eq "vhd") -OR ($vhdtype1 -eq "avhd"))
		{
		Add-Content $FileName "<td width='6%' BGCOLOR='#A00000' align=center>$vd1 GB</td>"
		Add-Content $FileName "<td width='6%' BGCOLOR='#A00000' align=center>$vdu1 GB</td>"
		}
	Else
		{
		Add-Content $FileName "<td width='6%'  align=center>$vd1 GB</td>"
		Add-Content $FileName "<td width='6%'  align=center>$vdu1 GB</td>"
		}
	Add-Content $FileName "<td width='5%' BGCOLOR='$vd2StorageHealth' align=center>$vd2Storage</td>"
	If (($vdtype2 -eq "vhd") -OR ($vhdtype2 -eq "avhd"))
		{
		Add-Content $FileName "<td width='6%' BGCOLOR='#A00000' align=center>$vd2 GB</td>"
		Add-Content $FileName "<td width='6%' BGCOLOR='#A00000' align=center>$vdu2 GB</td>"
		}
	Else
		{
		Add-Content $FileName "<td width='6%' align=center>$vd2 GB</td>"
		Add-Content $FileName "<td width='6%' align=center>$vdu2 GB</td>"
		}
	If ($vNetworkInterfaceType -eq "False")
		{
		Add-Content $FileName "<td width='6%' BGCOLOR='#F4FA58' align=center>Legacy</td>"
		}
	Else
		{
		Add-Content $FileName "<td width='6%' align=center>Synthetic</td>"
		}
	Add-Content $FileName "<td width='6%' align=center>$SSDate</td>"
	Add-Content $FileName "</tr>"
    }

Function fWriteHostInfo
	{ 
    	Param($FileName, $bgC, $sName, $tVMs, $pMem, $aMem, $allocStartup, $allocMaxmem, $aProc, $aCore, $aLogicalProc, $allocvProc, $hhealth)
	Add-Content $FileName "<tr bgcolor=$bgC>" 
	Add-Content $FileName "<td width='10%' BGCOLOR='$hhealth' align=center>$sName</td>"
	Add-Content $FileName "<td width='10%' BGCOLOR='$hhealth' align=center>$tVMs</td>" 
	Add-Content $FileName "<td width='10%' BGCOLOR='$hhealth' align=center>$pMem GB</td>" 
	Add-Content $FileName "<td width='10%' BGCOLOR='$hhealth' align=center>$aMem GB</td>"
	Add-Content $FileName "<td width='10%' align=center>$allocStartup GB</td>"
    	Add-Content $FileName "<td width='10%' align=center>$allocMaxmem GB</td>"
    	Add-Content $FileName "<td width='10%' align=center>$aProc</td>"
	Add-Content $FileName "<td width='10%' align=center>$aCore</td>"
	Add-Content $FileName "<td width='10%' align=center>$aLogicalProc</td>"
	Add-Content $FileName "<td width='10%' align=center>$allocvProc</td>"
	Add-Content $FileName "</tr>"
    }

Function fWriteStorageInfo
	{ 
    	Param($FileName, $volName, $volTotal, $volFree, $volFreePC, $VHDXAlloc, $VHDXActual, $OverProvisioned, $pcoverprov, $volhealthcode)
	Add-Content $FileName "<tr bgcolor=#77FF5B>"
	Add-Content $FileName "<td width='10%' BGCOLOR='$volhealthcode' align=center>$volName</td>"
	Add-Content $FileName "<td width='10%' BGCOLOR='$volhealthcode' align=center>$volTotal TB</td>"
	Add-Content $FileName "<td width='10%' BGCOLOR='$volhealthcode' align=center>$volFree GB</td>"
	Add-Content $FileName "<td width='10%' BGCOLOR='$volhealthcode' align=center>$volFreePC %</td>"
    	Add-Content $FileName "<td width='10%' align=center>$VHDXAlloc GB</td>"
    	Add-Content $FileName "<td width='10%' align=center>$VHDXActual GB</td>"
	Add-Content $FileName "<td width='10%' align=center>$OverProvisioned GB</td>"
	Add-Content $FileName "<td width='10%' align=center>$pcoverprov %</td>"
	Add-Content $FileName "</tr>"
    }
Function fCreateDashBoard
	{ 
    	Param($Name, $Type)
	If ($Type -match "Cluster")
			{
			Write-Host ("Fetching CSV Information from Cluster "+$Name) -foregroundcolor green
			[array] $CSVstorage = Get-ClusterSharedVolume -Cluster $Name| select -Expand SharedVolumeInfo |select FriendlyVolumeName, @{n="Name";e={($_.friendlyvolumename).TrimStart("C:\ClusterStorage\")}},  @{n="Capacity";e={$_.Partition.Size}}, @{n="UsedSpace";e={$_.Partition.UsedSpace}}, @{n="FreeSpace";e={$_.Partition.FreeSpace}}, @{n="FreeSpacePC";e={$_.Partition.PercentFree}}, @{n="VHDXAllocatedSpace";e={0}}, @{n="VHDXActualUsage";e={0}}, @{n="VolumeHealthCode";e={[int]"0"}}
			$CSVCount = $CSVStorage.length
			For ($h=0; $h -lt $CSVCount; $h++)
			{
			If (($CSVStorage.FreeSpacepc[$h] -le "10") -OR ($CSVStorage.FreeSpace[$h] -lt "53687091200")){[int] $CSVStorage[$h].VolumeHealthCode = "3"}
			ElseIf ((($CSVStorage.FreeSpacepc[$h] -le "20") -And ($CSVStorage.FreeSpace[$h] -gt 10)) -OR ($CSVStorage.FreeSpace[$h] -lt "107374182400")){[int] $CSVStorage[$h].VolumeHealthCode = "2"}
			ElseIf ((($CSVStorage.FreeSpacepc[$h] -le "30") -And ($CSVStorage.FreeSpace[$h] -gt 20)) -OR ($CSVStorage.FreeSpace[$h] -lt "214748364800")){[int] $CSVStorage[$h].VolumeHealthCode = 1}
			Else {$CSVStorage[$h].VolumeHealthCode = 0}
			}
		[array] $cNodes = Get-Cluster $Name|Get-ClusterNode |Where {$_.state -eq "Up"} |Select Name
		$nodecount = $cNodes.length
		[array] $hostDetails = Get-Cluster $Name|Get-ClusterNode|Select name, @{Label="TotalPhysicalMemory"; Expression={[int]""}}, @{Label="AvailablePhysicalMemory"; Expression={[int]""}}, @{Label="AvailablePhysicalMemoryPC"; Expression={[int]""}}, @{Label="TotalVMStartupRamAllocated"; Expression={[int]""}}, @{Label="TotalVMMaxMemoryAllocated"; Expression={[int]""}}, @{Label="Processors"; Expression={[int]""}}, @{Label="ProcessorCore"; Expression={[int]""}}, @{Label="LogicalProcessors"; Expression={[int]""}}, @{Label="vProcessors"; Expression={[int]""}}, @{Label="HostMemoryHealth"; Expression={[int]""}}, @{Label="VMs"; Expression={[int]""}}
		For ($i=0; $i -lt $nodecount; $i++)
    			{
			Write-Host ("Processing Hyper-V Host "+$cNodes.Name[$i]) -Foregroundcolor Green -BackgroundColor DARKGREEN
			$hDetails = Get-WmiObject -Class win32_OperatingSystem -ComputerName $cNodes.Name[$i] |Select FreePhysicalMemory, TotalVisibleMemorySize
			[int] $hostDetails[$i].TotalPhysicalMemory =  ((($hDetails).TotalVisibleMemorySize)/1048576)
			[int] $hostDetails[$i].AvailablePhysicalMemoryPC  =  (((($hDetails).FreePhysicalMemory)/(($hDetails).TotalVisibleMemorySize))*100)
			[int] $hostDetails[$i].AvailablePhysicalMemory =  ((($hDetails).FreePhysicalMemory)/1048576)
			If (($hostDetails[$i].AvailablePhysicalMemoryPC -le "10") -OR ($hostDetails[$i].AvailablePhysicalMemory -lt "5"))
				{
				[int] $hostDetails[$i].HostMemoryHealth = "3"
				}
			ElseIf ((($hostDetails[$i].AvailablePhysicalMemoryPC -le "20") -And ($hostDetails[$i].AvailablePhysicalMemoryPC -gt "10")) -OR (($hostDetails[$i].AvailablePhysicalMemory -lt "10") -And ($hostDetails[$i].AvailablePhysicalMemory -gt "5")))
				{
				[int] $hostDetails[$i].HostMemoryHealth = 2
				}
			ElseIf ((($hostDetails[$i].AvailablePhysicalMemoryPC -le "30") -And ($hostDetails[$i].AvailablePhysicalMemoryPC -gt "20")) -OR (($hostDetails[$i].AvailablePhysicalMemory -lt "20") -And ($hostDetails[$i].AvailablePhysicalMemory -gt "10")))
				{
				[int] $hostDetails[$i].HostMemoryHealth = 1
				}
			Else
				{
				[int] $hostDetails[$i].HostMemoryHealth = 0
				}
			$procDetails = Get-WmiObject -Class Win32_Processor -Computername $cNodes.Name[$i]
			$hostDetails[$i].Processors = ($procDetails.DeviceID).Count
			$hostDetails[$i].ProcessorCore = ($procDetails.numberofcores |Measure-Object -Sum).sum
			$hostDetails[$i].LogicalProcessors = ($procDetails.numberoflogicalprocessors |Measure-Object -Sum).sum
		}
		For ($i=0; $i -lt $nodecount; $i++)
    			{
			Write-Host ("Processing Virtual Servers on Hyper-V Host "+$cNodes.Name[$i]) -Foregroundcolor Green -BackgroundColor DARKGREEN
			[array] $vmList = Get-VM -Computer $cNodes.Name[$i] |Select Name
			$vmCount = $hostDetails[$i].VMs =  $vmList.count
			If ($vmCount -ge "1")
				{
				fWriteSubRowNodeName $ResultFile $cNodes.Name[$i] $hostDetails[$i].TotalPhysicalMemory $hostDetails[$i].AvailablePhysicalMemory $hostDetails[$i].AvailablePhysicalMemoryPC $WarningLevel[$hostDetails[$i].HostMemoryHealth]
				fWriteVMTableHeader $ResultFile
				For ($j=0; $j -lt $VMcount; $j++)
					{
					Write-Host ("Processing VM - "+$VMList[$j].name) -foregroundcolor Green
					$vmDetails = Get-VM -VMName $vmList[$j].name -Computer $cNodes.name[$i] |Select name, processorcount, MemoryDemand, VMID, IsClustered, @{Label="MinimumRam"; Expression={[System.Math]::Round(($_.MemoryMinimum/1073741824),2)}}, @{Label="MaximumRam"; Expression={[int]($_.MemoryMaximum/1073741824)}}, DynamicMemoryEnabled, @{Label="StartupRam"; Expression={[System.Math]::Round(($_.MemoryStartup/1073741824),2)}}, IntegrationServicesVersion, Uptime, ParentSnapshotId, IntegrationServicesState, ResourceMeteringEnabled
					$vmHealth = [int] ($hostDetails[$i]).HostMemoryHealth
					$hostDetails[$i].vProcessors += $vmDetails.processorcount
					If ($vmDetails.DynamicMemoryEnabled -eq "True")
						{
						$hostDetails[$i].TotalVMStartupRamAllocated += $vmDetails.StartupRam
						$hostDetails[$i].TotalVMMaxMemoryAllocated += $vmDetails.MaximumRam
						}
					Else
						{
						$hostDetails[$i].TotalVMStartupRamAllocated += $vmDetails.StartupRam
						$hostDetails[$i].TotalVMMaxMemoryAllocated += $vmDetails.StartupRam
						$VMDetails[0].MinimumRam = $vmDetails[0].MaximumRam = "DM Disabled"
						}
					If ($vmDetails.ResourceMeteringEnabled -eq "True")
						{
						[int] $AvgMemUsage = (((Measure-VM -ComputerName $cNodes.Name[$i]  -VMName $VMList.name[$j]).AvgRam)/1024)
						}
					Else 
						{
						[string] $AvgMemUsage = "NA"
						}
					If (Get-VHD -VMID $vmDetails.VMID -ComputerName $cNodes.Name[$i])
						{
						[array] $DiskDetails = Get-VHD -VMID $vmDetails.VMID -ComputerName $cNodes[$i].Name | Select path, @{Label="VolName"; Expression={""}}, @{Label="AllocatedSize"; Expression={$_.Size}}, @{Label="CurrentUsage"; Expression={$_.FileSize}}, @{Label="vDiskType"; Expression={""}}, @{Label="vDiskVolumeHealth"; Expression={""}}
						$vDiskcount = $DiskDetails.count
						For ($k = 0; $k -lt $vdiskcount; $k++)
							{		
								$VolInfo = (($DiskDetails[$k]).path).split('\')[2]
								$DiskDetails[$k].VolName = IF (($DiskDetails[$k]).Path -match "C:\\ClusterStorage") {($DiskDetails[$k].Path).Split("\")[2]} Else {($DiskDetails[$k].Path).Substring(0,2)}
								$DiskDetails[$k].VolName =  (Get-Culture).textinfo.totitlecase($DiskDetails[$k].VolName)
								$DiskDetails[$k].vDiskType = (($DiskDetails[$k]).path).split('.')[1]
								$VolIndex = [array]::indexof($CSVstorage.name,$DiskDetails[$k].VolName)
								$CSVStorage[$VolIndex].VHDXAllocatedSpace = $CSVStorage[$VolIndex].VHDXAllocatedSpace + $DiskDetails[$k].AllocatedSize
								$CSVStorage[$VolIndex].VHDXActualUsage = $CSVStorage[$VolIndex].VHDXActualUsage + $DiskDetails[$k].CurrentUsage
								$DiskDetails[$k].vDiskVolumeHealth = $CSVStorage.VolumeHealthCode[$VolIndex]
							}
						}
					$volHealthMax = $DiskDetails.vDiskVolumeHealth
					IF ($volHealthMax -gt $VMHealth) { $vmHealth = $volHealthMax }
					If ($VMDetails.ParentSnapshotId)
						{
						$SnapShotDate = ((Get-VMSnapshot -VMName $vmDetails.name -ComputerName $cNodes.Name[$i]).CreationTime |Sort-Object |Select-Object -First 1).ToShortDateString()
						}
					Else 
						{
						[String]$SnapShotDate = "NA"
						}
					 If ((Get-VMNetworkAdapter -VMName $vmDetails.Name -ComputerName $cNodes.Name[$i]).Count -eq "1")
						{
						$vNicType = (Get-VMNetworkAdapter -VMName $vmDetails.Name -ComputerName $cNodes.Name[$i]).IsLegacy
						}
					IF ($vDiskCount -gt "1") {
					fWriteVMInfo  $ResultFile $vmDetails.name $vmDetails.UpTime $vmDetails.IntegrationServicesVersion $vmDetails.IsClustered $vmDetails.ProcessorCount $vmDetails.StartupRam $vmDetails.MinimumRam $vmDetails.MaximumRam $AvgMemUsage $hostDetails[$i].HostMemoryHealth $DiskDetails[0].VolName ([int]($DiskDetails.AllocatedSize[0]/1073741824)) ([int]($DiskDetails.CurrentUsage[0]/1073741824)) $WarningLevel[$DiskDetails[0].vDiskVolumeHealth] $DiskDetails[0].vDiskType $DiskDetails[1].VolName  ([int]($DiskDetails.AllocatedSize[1]/1073741824)) ([int]($DiskDetails.CurrentUsage[1]/1073741824)) $WarningLevel[$DiskDetails[1].vDiskVolumeHealth]  $DiskDetails[1].vDiskType $vNicType  $SnapShotDate $vmDetails.IntegrationServicesState $WarningLevel[$vmHealth]
					}
					Else
					{
					fWriteVMInfo  $ResultFile $vmDetails.name $vmDetails.UpTime $vmDetails.IntegrationServicesVersion $vmDetails.IsClustered $vmDetails.ProcessorCount $vmDetails.StartupRam $vmDetails.MinimumRam $vmDetails.MaximumRam $AvgMemUsage $hostDetails[$i].HostMemoryHealth $DiskDetails[0].VolName ([int]($DiskDetails.AllocatedSize[0]/1073741824)) ([int]($DiskDetails.CurrentUsage[0]/1073741824)) $WarningLevel[$DiskDetails[0].vDiskVolumeHealth] $DiskDetails[0].vDiskType $DiskDetails[1].VolName  ([int]($DiskDetails.AllocatedSize[1]/1073741824)) ([int]($DiskDetails.CurrentUsage[1]/1073741824)) $WarningLevel[0]  $DiskDetails[1].vDiskType $vNicType  $SnapShotDate $vmDetails.IntegrationServicesState $WarningLevel[$vmHealth]
					}
					Write-Host "Finished processing VM  - " $vmDetails.name -ForegroundColor Yellow -BackgroundColor DarkGreen
					}
				Write-Host "Finished Processing  VMs on Hyper-V Cluster Node - "  $cNodes.Name[$i] -ForegroundColor white -BackgroundColor BLUE
        		Add-Content $ResultFile "</table>"
				}
			Else
				{
				Write-Host "NO  VMs Found on Hyper-V Server - " $cNodes.Name[$i] -Foregroundcolor Black -backgroundcolor DarkRed
				}
		}
		fWriteSubHeadingHostDetails $ResultFile $Name
		fWriteHostStatusTableHeader $ResultFile
		For ($m = 0; $m -lt $nodecount; $m++)
			{
			fWriteHostInfo $ResultFile "#77FF5B" $hostDetails[$m].name $hostDetails[$m].VMs $hostDetails[$m].TotalPhysicalMemory $hostDetails[$m].AvailablePhysicalMemory $hostDetails[$m].TotalVMStartupRamAllocated $hostDetails[$m].TotalVMMaxMemoryAllocated $hostDetails[$m].processors $hostDetails[$m].ProcessorCore $hostDetails[$m].LogicalProcessors $hostDetails[$m].vProcessors $WarningLevel[($hostDetails[$m]).HostMemoryHealth]
			}
		fWriteHostInfo $ResultFile "#FA58D0" "TOTAL" ($hostDetails.VMs|Measure-Object -Sum).sum ($hostDetails.TotalPhysicalMemory|Measure-Object -Sum).sum ($hostDetails.AvailablePhysicalMemory|Measure-Object -Sum).sum ($hostDetails.TotalVMStartupRamAllocated|Measure-Object -Sum).sum ($hostDetails.TotalVMMaxMemoryAllocated|Measure-Object -Sum).sum ($hostDetails.processors|Measure-Object -Sum).sum ($hostDetails.ProcessorCore |Measure-Object -Sum).sum ($hostDetails.LogicalProcessors|Measure-Object -Sum).sum ($hostDetails.vProcessors |Measure-Object -Sum).sum
		Add-Content $ResultFile "</table>"
		fWriteSubHeadingCSVDetails $ResultFile $Name
		Write-Host "Processing CSV Storage details $name " -ForegroundColor "Yellow"
		fWriteStorageTableHeader $ResultFile
		$len = $CSVStorage.count
		For ($z=0; $z -lt $len; $z++)
			{ 
			$csvName = $CSVStorage[$z].name
			$csvTotalSize = [System.Math]::Round(($CSVStorage[$z].Capacity/1099511627776),2)
			[int] $csvFree = (($CSVStorage[$z].Freespace)/1073741824)
			[int] $csvFreePC = $CSVStorage[$z].Freespacepc
			[int] $vhdxAlloc = (($CSVStorage[$z].VHDXAllocatedSpace)/1073741824)
			[int] $vhdxActual = (($CSVStorage[$z].VHDXActualUsage)/1073741824)
			[int] $vhdxOP = (($CSVStorage[$z].VHDXAllocatedSpace - $CSVStorage[$z].Capacity)/1073741824)
			[int] $pcOP = (($CSVStorage[$z].VHDXAllocatedSpace * 100) /$CSVStorage[$z].Capacity)
			If ($vhdxop -le "0")
				{
				[string] $vhdxop = "Not Over Provisioned"
				[string] $pcop = "NA"
				}
			fWriteStorageInfo $ResultFile $csvName $csvTotalSize $csvFree $csvFreePC $vhdxAlloc $vhdxActual $vhdxOP $pcOP $WarningLevel[($CSVStorage[$z]).VolumeHealthCode]
			}
		Add-Content $ResultFile "</table>"
		$date = ( get-date ).ToString('yyyy/MM/dd') 
		Write-Host ("Finished Processing Storage for  Cluster $name") -ForegroundColor BLACK -BackgroundColor CYAN
	}
	ElseIf ($type -match "StandAlone")
	{
		Write-Host ("Processing Storage Information on StandAlone Hyper-V Host " +$Name) -Foregroundcolor Green
			[array] $LocalStorage = Get-WmiObject Win32_LogicalDisk -filter "DriveType=3" -computer $Name | Select DeviceID, Size, FreeSpace, @{n="FreeSpacePC";e={[int]($_.FreeSpace/$_.Size*100)}}, @{n="VHDXAllocatedSpace";e={0}}, @{n="VHDXActualUsage";e={0}}, @{n="VolumeHealthCode";e={[int]"0"}}
			$IsClustered = "NA"
			$LocalStorageCount = ($LocalStorage).count
			For ($k=0; $k -lt $LocalStorageCount; $k++)
			{
			If (($LocalStorage.FreeSpacePC[$k] -le "10") -OR ($LocalStorage.FreeSpace[$k] -lt "53687091200")){[int] $LocalStorage[$k].VolumeHealthCode = "3"}
			ElseIf ((($LocalStorage.FreeSpacePC[$k] -le "20") -And ($LocalStorage.FreeSpace[$k] -gt 10)) -OR ($LocalStorage.FreeSpace[$k] -lt "107374182400")){[int] $LocalStorage[$k].VolumeHealthCode = "2"}
			ElseIf ((($LocalStorage.FreeSpacePC[$k] -le "30") -And ($LocalStorage.FreeSpace[$k] -gt 20)) -OR ($LocalStorage.FreeSpace[$k] -lt "214748364800")){[int] $LocalStorage[$k].VolumeHealthCode = 1}
			Else {$LocalStorage[$k].VolumeHealthCode = 0}
			}

		Write-Host "Processing Physical Host Details " $Name -Foregroundcolor Green -BackgroundColor DARKGREEN
			[array] $hostDetails = Get-WmiObject -Class win32_OperatingSystem -ComputerName $Name |Select @{Label="TotalPhysicalMemory"; Expression={[int]($_.TotalVisibleMemorySize/1048576)}}, @{Label="AvailableMemory"; Expression={[int]($_.FreePhysicalMemory/1048576)}}, @{Label="AvailablePhysicalMemoryPC"; Expression={[int](($_.FreePhysicalMemory/$_.TotalVisibleMemorySize)*100)}}
			[Int] $TotalVMMaxMemoryAllocated = [int] $TotalVMStartupRamAllocated = [int] $vProcessors = 0
			If (($hostDetails.AvailablePhysicalMemoryPC -le "10") -OR ($hostDetails.AvailableMemory -lt "5"))
				{
				[int] $HostMemoryHealth = 3
				}
			ElseIf ((($hostDetails.AvailablePhysicalMemoryPC -le "20") -And ($hostDetails.AvailablePhysicalMemoryPC -gt "10")) -OR (($hostDetails.AvailableMemory -lt "10") -And ($hostDetails.AvailableMemory -gt "5")))
				{
				[int] $HostMemoryHealth = 2
				}
			ElseIf ((($hostDetails.AvailablePhysicalMemoryPC -le "70") -And ($hostDetails.AvailablePhysicalMemoryPC -gt "20")) -OR (($hostDetails.AvailableMemory -lt "20") -And ($hostDetails.AvailableMemory -gt "10")))
				{
				[int] $HostMemoryHealth = 1
				}
			Else
				{
				[int] $HostMemoryHealth = 0
				}
			$procDetails = Get-WmiObject -Class Win32_Processor -Computername $Name
			$Processors = ($procDetails.DeviceID).Count
			$ProcessorCore = ($procDetails.numberofcores |Measure-Object -Sum).sum
			$LogicalProcessors = ($procDetails.numberoflogicalprocessors |Measure-Object -Sum).sum
			[array] $vmList = Get-VM -Computer $Name
			$vmCount = ($vmList.Name).count
			If ($vmCount -gt "0")
				{
				Write-Host "Processing Virtual Servers on Hyper-V Standalone Host - " $Name -Foregroundcolor Green -BackgroundColor DARKGREEN
				fWriteSubRowNodeName $ResultFile $Name $hostDetails.TotalPhysicalMemory $hostDetails.AvailablePhysicalMemory $hostDetails.AvailablePhysicalMemoryPC $WarningLevel[$HostMemoryHealth]
				fWriteVMTableHeader $ResultFile
				For ($j=0; $j -lt $vmCount; $j++)
					{
					Write-Host "Processing VM - " $VMList[$j].name -ForegroundColor Yellow
					[array]$vmDetails = Get-VM -VMName $vmList[$j].name -Computer $Name |Select name, processorcount, MemoryDemand, VMID, @{Label="MinimumRam"; Expression={[int]($_.MemoryMinimum/1073741824)}}, @{Label="MaximumRam"; Expression={[int]($_.MemoryMaximum/1073741824)}}, DynamicMemoryEnabled, @{Label="StartupRam"; Expression={[int]($_.MemoryStartup/1073741824)}}, IntegrationServicesVersion, Uptime, ParentSnapshotId, IntegrationServicesState, ResourceMeteringEnabled
					If ($vmDetails.DynamicMemoryEnabled -eq "True")
						{
						$TotalVMStartupRamAllocated += $vmDetails.StartupRam
						$TotalVMMaxMemoryAllocated += $vmDetails.MaximumRam
						}
					Else
						{
						$TotalVMStartupRamAllocated += $vmDetails.StartupRam
						$TotalVMMaxMemoryAllocated += $vmDetails.StartupRam
						$VMDetails[0].MinimumRam = $vmDetails[0].MaximumRam = "DM Disabled"
						}
					If ($vmDetails.ResourceMeteringEnabled -eq "True")
						{
						[int] $AvgMemUsage = (((Measure-VM -ComputerName $Name  -VMName $VMList.name[$j]).AvgRam)/1024)
						}
					ElseIf ($vmDetails.ResourceMeteringEnabled -match "False")
						{
						Write-Host "Resource Metering is disabled - Skipping Avg Memory on " $VMList[$j].name -ForegroundColor Red
						[string] $AvgMemUsage = "NA"
						}
					If (Get-VHD -VMID $vmDetails.VMID -ComputerName $Name -erroraction SilentlyContinue)
						{
						[array] $DiskDetails = Get-VHD -VMID $vmDetails.VMID -ComputerName $Name | Select path, @{Label="VolName"; Expression={""}}, @{Label="AllocatedSize"; Expression={$_.Size}}, @{Label="CurrentUsage"; Expression={$_.FileSize}},  @{Label="vDiskType"; Expression={""}}, @{Label="vDiskVolumeHealth"; Expression={""}}
						$vDiskcount = $DiskDetails.count
						For ($k = 0; $k -lt $vdiskcount; $k++)
							{
							$DiskDetails[$k].VolName = ($DiskDetails[$k].Path).Substring(0,2)
							$DiskDetails[$k].VolName =  (Get-Culture).textinfo.totitlecase($DiskDetails[$k].VolName)		
							$VolInfo = (($DiskDetails[$k]).path).Substring(0,2)
							$VolIndex = [array]::indexof($localStorage.DeviceID,$VolInfo)
							$DiskDetails[$k].vDiskType = (($DiskDetails[$k]).path).split('.')[1]
							$localStorage[$VolIndex].VHDXAllocatedSpace = $localStorage[$VolIndex].VHDXAllocatedSpace + $DiskDetails[$k].AllocatedSize
							$localStorage[$VolIndex].VHDXActualUsage = $localStorage[$VolIndex].VHDXActualUsage + $DiskDetails[$k].CurrentUsage
							$DiskDetails[$k].vDiskVolumeHealth = $localStorage[$VolIndex].VolumeHealthCode
							}
						}
					If ($VMDetails.ParentSnapshotId)
						{
						$SnapShotDate = ((Get-VMSnapshot -VMName $vmDetails.name -ComputerName $Name).CreationTime |Sort-Object |Select-Object -First 1).ToShortDateString()
						}
					Else 
						{
						[String]$SnapShotDate = "NA"
						}
					 If ((Get-VMNetworkAdapter -VMName $vmDetails.Name -ComputerName $Name).Count -eq "1")
						{
						$vNicType = (Get-VMNetworkAdapter -VMName $vmDetails.Name -ComputerName $Name).IsLegacy
						}
					If ($vDiskcount -gt "1")
					{
					fWriteVMInfo  $ResultFile $vmDetails.name $vmDetails.UpTime $vmDetails.IntegrationServicesVersion "NA" $vmDetails.ProcessorCount $vmDetails.StartupRam $vmDetails.MinimumRam $vmDetails.MaximumRam $AvgMemUsage $HostMemoryHealth $DiskDetails[0].VolName ([int]($DiskDetails.AllocatedSize[0]/1073741824)) ([int]($DiskDetails.CurrentUsage[0]/1073741824)) $WarningLevel[$DiskDetails[0].vDiskVolumeHealth] $DiskDetails[0].vDiskType $DiskDetails[1].VolName ([int]($DiskDetails.AllocatedSize[1]/1073741824)) ([int]($DiskDetails.CurrentUsage[1]/1073741824)) $WarningLevel[$DiskDetails[1].vDiskVolumeHealth] $DiskDetails[1].vDiskType $vNicType $SnapShotDate $vmDetails.IntegrationServicesState $WarningLevel[$vmHealth]
					}
					Else
					{
					fWriteVMInfo  $ResultFile $vmDetails.name $vmDetails.UpTime $vmDetails.IntegrationServicesVersion "NA" $vmDetails.ProcessorCount $vmDetails.StartupRam $vmDetails.MinimumRam $vmDetails.MaximumRam $AvgMemUsage $HostMemoryHealth $DiskDetails[0].VolName ([int]($DiskDetails.AllocatedSize[0]/1073741824)) ([int]($DiskDetails.CurrentUsage[0]/1073741824)) $WarningLevel[$DiskDetails[0].vDiskVolumeHealth] $DiskDetails[0].vDiskType $null $null $null $null $null $vNicType $SnapShotDate $vmDetails.IntegrationServicesState
					}
					Write-Host "Finished Processing " $VMDetails.name -ForegroundColor Yellow -BackgroundColor DarkGreen
					}
				Write-Host "Finished Processing  VMs on Hyper-V Server " $Name -ForegroundColor Yellow -BackgroundColor DarkGreen
        			Add-Content $ResultFile "</table>"
				}
			Else
				{
				Write-Host "NO  VMs Found on Hyper-V Server $Name" -ForegroundColor Black -BackgroundColor Red
				}
		fWriteSubHeadingHostDetails $ResultFile $Name
		fWriteHostStatusTableHeader $ResultFile
		fWriteHostInfo $ResultFile "#77FF5B" $name $vmCount $hostDetails.TotalPhysicalMemory $hostDetails.AvailableMemory $TotalVMStartupRamAllocated $TotalVMMaxMemoryAllocated $processors $ProcessorCore $LogicalProcessors $vProcessors $WarningLevel[$HostMemoryHealth]
		Add-Content $ResultFile "</table>"
		fWriteSubHeadingCSVDetails $ResultFile $Name
		fWriteStorageTableHeader $ResultFile
		$len = $localStorage.count
		For ($z=0; $z -lt $len; $z++)
			{ 
			$localDiskName = $localStorage[$z].deviceID
			$localDiskTotalSize = [System.Math]::Round((($localStorage[$z].Size)/1099511627776),2)
			[int] $localDiskFree = [System.Math]::Round((($localStorage[$z].Freespace)/1073741824),2)
			[int] $localDiskFreePC = (($localStorage[$z].Freespace * 100) / $localStorage[$z].Size)
			[int] $vhdxAlloc = (($localStorage[$z].VHDXAllocatedSpace)/1073741824)
			[int] $vhdxActual = (($localStorage[$z].VHDXActualUsage)/1073741824)
			[int] $vhdxOP = (($localStorage[$z].VHDXAllocatedSpace - $localStorage[$z].Size)/1073741824)
			[int] $pcOP = (($localStorage[$z].VHDXAllocatedSpace*100)/$localStorage[$z].Size)
			If ($vhdxop -le "0")
				{
				[string] $vhdxop = "Not Over Provisioned"
				[string] $pcop = "NA"
				}
			fWriteStorageInfo $ResultFile $localDiskName $localDiskTotalSize $localDiskFree $localDiskFreePC  $vhdxAlloc $vhdxActual $vhdxOP $pcOP
			}
		Add-Content $ResultFile "</table>"
		$date = ( get-date ).ToString('yyyy/MM/dd') 
		Write-Host ("Finished Processing Storage for  Hyper-V Standalone Host $name ") -ForegroundColor DarkRed -BackgroundColor DarkGreen
		}
	}
fWriteHtmlHeader $ResultFile
Write-Host "Fetching Input File" -foregroundColor Green
If (Test-Path .\servers.txt)
	{
	$sList = (Get-Content .\servers.txt)
	$sCount = $sList.Count
	If ($sCount -eq "0")
	{
	Write-Host "Servers.txt Empty. Please check if the input file exists and populated with server names"
	}
	If ($sCount -gt "0")
	{
	Write-Host "Identified Servers.txt. Checking if Servers/Clusters are Available"
	If ($sCount -eq "1")
		{
		[array] $sList = $sList
		}
	For ($a = 0; $a -lt $sCount; $a++)
	{
		Write-Host ("Checking "+$sList[$a])
		If (Get-Cluster $sList[$a] -erroraction SilentlyContinue)
		{
		Write-Host ("Identified cluster "+$sList[$a]) -Background White -Foreground Black
		fCreateDashBoard $sList[$a] "Cluster"
		}
		ElseIf ((get-vmhost $sList[$a]) -And (!(Get-Service -name clussvc -ComputerName $sList[$a] -erroraction SilentlyContinue)))
		{
		Write-Host ("Identified Standalone HyperV Server - "+$sList[$a]) -Background White -Foreground Black
		fCreateDashBoard $sList[$a] "StandAlone"
		}
		Else
		{
		Write-Host "Not able to identify the server as a StandAlone host or Cluster $sList[$a] " -Background White -Foreground Black
		}
	}
	}
	}
ElseIf ((Get-Cluster) -Or (Get-VMHost))
	{
	Write-Host "Checking local host for Hyper-V Role"
	IF (Get-Cluster)
		{
		$sList = (Get-Cluster).name
		fCreateDashBoard $sList "Cluster"
		}
	ElseIf (Get-VMHost)
		{
		$sList = (Get-VmHost).name
		fCreateDashBoard $sList "StandAlone"
		}
	}
ElseIf ((!(Get-Cluster)) -And (!(Get-VMHost)) -And (!(Test-Path Servers.txt)))
	{
	Write-Host "Please provide input file or run the script from a Hyper-V Standlone Server / Hyper-V Cluster node" -ForegroundColor BLACK -BackgroundColor "Red"
	Write-Host "#########  EXITING SCRIPT ########### " -ForegroundColor Yellow -BackgroundColor "Red"
	}
fWriteHtmlFooter $ResultFile
Send-MailMessage -To email@yourdomain.com -From HyperV-Dashboard@InsideVirtualization.com -Body (Get-Content HyperV-VM-Report.htm |Out-String) -SmtpServer Smtp.yourdomain.com -Subject "HyperV Health Report" -BodyAsHtml