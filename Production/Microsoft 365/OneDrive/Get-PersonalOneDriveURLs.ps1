#################################################################
#																                                #
#	    Script to gather OneDrive Personal URLs from Office365		#
#																                                #
#	    Written by: AHEAD, Inc. 							                    #
#	    Creation Date: 3/22/2022						                    	#
#																                                #
#																                                #
#################################################################


# Set temp path and create if not exists.
Write-Host "Checking for C:\Temp directory" -ForegroundColor Yellow
$TempPath = "C:\Temp\"
	if (!(Test-Path $TempPath))
	{
	New-Item -itemType Directory -Path "C:\Temp\"
	}
	else
	{
	write-host "Folder already exists" -ForegroundColor Green
	}


# Start Transcript
Start-Transcript -Path "C:\Temp\Get-OneDriveURLs.txt" -Append


# Define Tenant Admin URL. Specify the full admin Url <https://contoso-admin.sharepoint.com>
$TenantUrl = $(Write-Host "Enter the SharePoint admin center URL:  " -ForegroundColor Magenta -NoNewLine) + $(Read-Host) 


# Connect to SharePoint Online Service
Connect-SPOService -Url $TenantUrl


# Get Personal OneDrive Site URLs, Export to CSV and Log
Get-SPOSite -IncludePersonalSite $true -Limit all -Filter "Url -like '-my.sharepoint.com/personal/'" | Select Url, Title, Owner, StorageQuota, StorageQuotaWarningLevel, ResourceQuota, ResourceQuotaWarningLevel, Template, Status | Export-CSV -Path C:\Temp\Get-OneDriveURLs.csv -nti
Write-Host "Done! Exported to $TempPath as CSV." -ForegroundColor Green


# Disconnect SPO Session
Write-Host "Disconnecting SPOnline Session." -ForegroundColor Yellow
Disconnect-SPOService


# Stop Transcript
Stop-Transcript
