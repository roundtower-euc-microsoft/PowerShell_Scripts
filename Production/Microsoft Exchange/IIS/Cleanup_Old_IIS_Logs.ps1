Import-Module WebAdministration  
  
#Maximum age in days of files to be deleted  
$logfileMaxAge = 2  
foreach($website in $(Get-Website))  
{  
    #Get log folder for current website  
    $folder="$($website.logFile.directory)\W3SVC$($website.id)".replace("%SystemDrive%",$env:SystemDrive)  
    #Get all log files in the folder  
    $files = Get-ChildItem $folder -Filter *.log  
  
    foreach($file in $files){  
        if($file.LastWriteTime -lt (Get-Date).AddDays(-1*$logfileMaxAge)){  
            #Remove fie older than logfileMaxAge days  
            Remove-Item $file.FullName  
  
        }  
    }  
}  