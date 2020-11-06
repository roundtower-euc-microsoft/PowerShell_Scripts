<#
.SYNOPSIS
Script to parse through the smtp protocol logs and list out the unique IPs and DNS names
.DESCRIPTION 
Outputs a list of the IPs and DNS names of the Unique IPs found in the SMTP Protocol Logs
.EXAMPLE
.\GetSMTPLogUnique.ps1 -dir "F:\Program Files\Microsoft\Exchange Server\V14\TransportRoles\Logs\ProtocolLog\SmtpReceive"
.LINK
https://github.com/mikecessna/ExchangeScripts/blob/master/GetSMTPLogUniqueIPs.ps1
.NOTES
Written By: Mike Cessna
Change Log
V1.00, 6/1/2014 Initial Version
#>
[CmdletBinding()]
Param(
	[Parameter(Mandatory=$true)][string] $dir)
#if you want to exclude IPs (like other Exchange Servers in your org) put them here.
$Exclude=@("10.12.34.10","10.12.34.11","10.10.253.196","10.10.254.139","10.10.254.143","10.13.34.10","10.10.254.20")

#Get all the logs in the specified Dir
$files=Get-ChildItem $dir -Filter *.log
#loop through the files unless the Dir is empty
if($files -ne $null){
	$logs=@()
	$output=@()
	foreach ($file in $files) {
        #grab the file, skip any lines that begin with a #
        #split the log by comma and take column 5 (default for client IP
        #then split the filed by : and take jsut the first element, this is the IP of the client
	    $logs+=get-content $file.FullName | ?{$_ -notmatch "^#"} | % {$_.Split(",")[5]} | %{$_.Split(":")[0]}
	}
    #sort the logs and only keep the Unique IPs
	$logs= $logs | sort-object | get-unique
    #Loop through the IPs and resolve the name from DNS
    #and set up an object for output
	foreach($log in $logs){
	    if($Exclude -notcontains $log){
	        $dns=$null
	        Write-Verbose "Working on $log"
	        $objlog = new-object system.object
	        $objlog | add-member -type NoteProperty -name IP -value $log
	        $dns=[System.Net.Dns]::GetHostEntry($log).HostName
            #GetHostEntry will return the ip back if it can't resolve the IP to a name
            #so use the resolved name if you don't get the IP back
	        if($dns -ne $log){
	            $objlog | add-member -type NoteProperty -name DNS-Name -value $dns
	        } else {
                #if it doesn't resolve use "Unknown" for the name
	            $dns="Unknown"
	            $objlog | add-member -type NoteProperty -name DNS-Name -value $dns
	        }
	        Write-Verbose "Got this from DNS $dns"
            #push the info into our output var
	        $output+=$objlog
	    }
	}
	$output
} else {
    #if the DIR is empty tell me
	Write-Host "You log Directory ($dir) appears to be empty of log files."
	Write-Host "Check your path and try again."
	Write-Host "Also ensure that your Exchange server Protocol Logging is enabled."
}