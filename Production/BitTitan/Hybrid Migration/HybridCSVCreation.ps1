# +---------------------------------------------------------------------------
# | File : HybridCSVCreation.ps1
# | Version : 1.5
# | Description : This script will create the CSV file for importing into MigrationWiz
# | Usage : .\HybridCSVCreation.ps1
# | Author: Mark Rochester - markr@bittitan.com
# +-------------------------------------------------------------------------------
$ErrorActionPreference = "SilentlyContinue"

write-host "BitTitan Exchange Collector - (c) 23 September 2020"
write-host "------------------------------------------"
write-host "Adding Exchange Management Snap In"

# Add the Exchange PowerShell snap-in to the current console
$exchangeSnapins = Get-PSSnapin | Where-Object {$_.Name.Contains("Microsoft.Exchange.Management.PowerShell")}
if (-not $exchangeSnapins) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.*
}

$CSVfile = Read-Host “Enter the Path of CSV file (Eg. C:\Report.csv)”
$TenantName = Read-Host "Enter the Tenant name that preceeds the 'onmicrosoft.com' namespace "
write-host "Retrieving Mailboxes"
$AllMailbox = Get-mailbox -resultsize Unlimited | ? {$_.RecipientTypeDetails -ne "DiscoveryMailbox"}
$mbxcount = $AllMailbox.count
$mbxprocess = 1
$output = @()

write-host "Found $mbxcount Mailboxes"

Foreach($mbx in $AllMailbox) {

Write-Progress -Activity “Report in Progress” -Status “Processing $mbxprocess of $mbxcount – $($mbx.displayname)” -percentcomplete ($mbxprocess/$mbxcount*100)

    $userObj = New-Object PSObject
    $userObj | Add-Member NoteProperty -Name "SMTPAddress” -Value $mbx.PrimarySmtpAddress
    $userObj | Add-Member NoteProperty -Name "UserPrincipalName" -Value $mbx.userprincipalname
    $userObj | Add-Member NoteProperty -Name "RecipientTypeDetails" -Value $mbx.recipienttypedetails
    $userObj | Add-Member NoteProperty -Name "GUID" -Value $mbx.exchangeguid


    # Get the organizatinal unit
	$organizationalUnit = $mbx.OrganizationalUnit

	# Remove the domain name in the organizational unit
	if ($organizationalUnit -and $organizationalUnit.IndexOf("/") -ge 0) {
		$organizationalUnit = [string]$organizationalUnit.Substring(($organizationalUnit.IndexOf("/") + 1))
	}

    # Calculate the Mailbox Stats
    $stats = get-mailboxstatistics -Identity $mbx.DistinguishedName
    $userObj | Add-Member NoteProperty -Name "MailboxItemsCount" -Value $stats.itemcount
    
    [double]$mbsize = ($stats.TotalItemSize -split "\s",3)[0]
    $mbtype = ($stats.totalitemsize -split "\s",3)[1]

    if ($mbtype -eq "B") { $mbsize = [math]::round($mbsize / 1024) }
    if ($mbtype -eq "KB") { $mbsize = [math]::round($mbsize) }
    if ($mbtype -eq "MB") { $mbsize = [math]::round($mbsize * 1024) }
    if ($mbtype -eq "GB") { $mbsize = [math]::round($mbsize * 1024 * 1024) }
    if ($mbtype -eq "TB") { $mbsize = [math]::round($mbsize * 1024 * 1024 * 1024) }
            
    $userObj | Add-Member NoteProperty -Name "MailboxSizeinKB" -Value $mbsize
    

	# Add the organizational unit
	$userObj | Add-Member NoteProperty -Name "OrganizationalUnit" -Value $organizationalUnit.replace(",","|") -Force
	# Add the guid
	$userObj | Add-Member NoteProperty -Name “BatchName” -Value ""

    # Check if user has appropriate proxy address
			$proxyaddress = $("smtp$($mailbox.alias)@$TenantName.mail.onmicrosoft.com")

			# Add Proxy address flag based on proxy address availability
			if (!($mbx.emailaddresses -contains $proxyaddress)) {
				$userObj | Add-Member NoteProperty -Name "ProxyAddressAvailable" -Value "False"
			} else {
				$userObj | Add-Member NoteProperty -Name "ProxyAddressAvailable" -Value "True"
			}

			# Get the CAS mailbox properties for the specified mailbox
			$casMailbox = Get-CASMailbox $mbx.Identity
    
			# Add the imap, pop, owa and active sync
			if ($casMailbox) {
				$userObj | Add-Member NoteProperty -Name "IMAP" -Value $([string]$casMailbox.ImapEnabled)
				$userObj | Add-Member NoteProperty -Name "POP" -Value $([string]$casMailbox.PopEnabled)
				$userObj | Add-Member NoteProperty -Name "OWA" -Value $([string]$casMailbox.OWAEnabled)
				$userObj | Add-Member NoteProperty -Name "ActiveSync" -Value $([string]$casMailbox.ActiveSyncEnabled)
			}
			else {
				$userObj | Add-Member NoteProperty -Name "IMAP" -Value ""
				$userObj | Add-Member NoteProperty -Name "POP" -Value ""
				$userObj | Add-Member NoteProperty -Name "OWA" -Value ""
				$userObj | Add-Member NoteProperty -Name "ActiveSync" -Value ""

				# Update status and status message
				$status = "Failure"
				$statusMessage = "Unable to retrieve CAS properties for {$mbx.Guid.ToString()}"
			}

        #Get Last Logon Time
        $userObj | Add-Member NoteProperty -Name "LastLogonTime" -Value $stats.lastlogontime

    
		# Get delegates information for the specified mailbox
		
        $delegates = Get-MailboxPermission $mbx.Identity | Where {$_.User.ToString() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false} | Select User | ConvertTo-Csv -NoTypeInformation | ForEach-Object {$_ -replace '"',''} | select -Skip 1 | Get-User | Select Identity | ConvertTo-Csv -NoTypeInformation | ForEach-Object {$_ -replace '"',''}
		
			# Check if delegates are not null
			if ($delegates) {
				# Skip the header line
				$delegates = $delegates | Select -Skip 1

				# Define delegatesSmtp list
				$delegatesSmtp = New-Object System.Collections.Generic.List[string]
			
				# Iterate through the list of delegates to get smtp of each delegate
				foreach ($delegate in $delegates) {
					# Get smtp for the specified delegate
					$delegateSmtp = Get-Mailbox $delegate | Select PrimarySmtpAddress

					# Check mailbox is not null
					if ($delegateSmtp)
					{
						$delegateSmtpStr = [string]$delegateSmtp.PrimarySmtpAddress
						if (-Not $delegateSmtpStr.Contains(",")) {
							# Add smtp to the list
							$delegatesSmtp.Add($delegateSmtpStr)
						} 
					}
				}
            
				# Concatenate all the delegates with | delimiter
				$delegatesSmtpString = $delegatesSmtp -Join '|'

				# Add the Delegates property
				$userObj | Add-Member NoteProperty -Name "Delegates" -Value $delegatesSmtpString -force
			} else {
				# Add the Delegates property
				$userObj | Add-Member NoteProperty -Name "Delegates" -Value "" -force
			}
        	$userObj | Add-Member NoteProperty -Name "Identity" -Value "" -Force

	
    $output += $userObj
    $mbxprocess ++

}


$output |export-csv -Path $CSVfile -NoTypeInformation

(Get-Content $CSVFile) -replace '(?m)"([^,]*?)"(?=,|$)', '$1' | Set-Content $CSVFile

write-host "Data Collected and stored in CSV"
