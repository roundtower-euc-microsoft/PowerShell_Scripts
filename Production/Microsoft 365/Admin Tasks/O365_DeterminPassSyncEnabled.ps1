<#
Description:
This script queries SQL to see if Password Hash Sync is enabled.
 
October 22 2013 (v2)
Mike Crowley
 
http://mikecrowley.us
 
Known Issues:
1) All commands, including SQL queries run as the local user. This may cause issues on locked-down SQL deployments.
2) For remote SQL installations, the SQL PowerShell module must be installed on the dirsync server.
 (http://technet.microsoft.com/en-us/library/hh231683.aspx)
3) Assumes Dirsync version 6385.0012 or later.
 
#>

#Console Prep
cls
Write-Host "Please wait..." -F Yellow

#Import and check for SQL Module
ipmo SQLps
cls
if ((gmo sqlps) -eq $null) {
 Write-host "The SQL PowerShell Module Is Not loaded." -F Magenta
 Write-host "Download and retry. http://technet.microsoft.com/en-us/library/hh231683.aspx" -F Magenta
 Write-Host
 Write-Host "Quitting..." -F Magenta; sleep 2; Write-host; break
 }

#Learn SQL Instance
$SQLServer = (gp 'HKLM:SYSTEM\CurrentControlSet\services\FIMSynchronizationService\Parameters').Server
If ($SQLServer.Length -eq '0') {$SQLServer = $env:computername}
$SQLInstance = (gp 'HKLM:SYSTEM\CurrentControlSet\services\FIMSynchronizationService\Parameters').SQLInstance
$MSOLInstance = ($SQLServer + "\" + $SQLInstance)

#Query for Password Hash Sync
[xml]$ADMAxml = Invoke-Sqlcmd -ServerInstance $MSOLInstance -Query "SELECT [ma_id] ,[ma_name] ,[private_configuration_xml] FROM [FIMSynchronizationService].[dbo].[mms_management_agent]" | ? {$_.ma_name -eq 'Active Directory Connector'} | select -Expand private_configuration_xml
If ((Select-Xml -XML $ADMAxml -XPath "/adma-configuration/password-hash-sync-config/enabled" | select -expand node).'#text' -eq '1') {Write-Host "Password Hash Sync is Enabled." -Fore Cyan}
If ((Select-Xml -XML $ADMAxml -XPath "/adma-configuration/password-hash-sync-config/enabled" | select -expand node).'#text' -eq '0') {Write-Host "Password Hash Sync NOT Enabled." -Fore Cyan}

Write-Host

#End of Script