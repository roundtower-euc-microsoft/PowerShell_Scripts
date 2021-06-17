Param(
    [string]$IMCEAEX
)

Write-Host ""
Write-Host -ForegroundColor Magenta "**********************************************************"
Write-Host -ForegroundColor Magenta "****** Convert your IMCEAEX NDR to an X.500 address ******"
Write-Host -ForegroundColor Magenta "**********************************************************"

If ($IMCEAEX -eq "") {
    Write-Host -ForegroundColor Yellow -NoNewline "`nPaste your IMCEAEX NDR string here: "
    $IMCEAEX = Read-Host
}

If($IMCEAEX.Substring(0,7) -ne "IMCEAEX") {
    Write-Host -ForegroundColor Red "`nSorry, your IMCEAEX string must begin with IMCEAEX`n" 
} Else {
    $X500 = $IMCEAEX.Replace("IMCEAEX-","X500:").Replace("_","/").Replace("+20"," ").Replace("+28","(").Replace("+29",")").Replace("+2E",".").Replace("%3D","=").Split("@")[0]
    Write-Host 
    Write-Host -ForegroundColor DarkCyan "Your converted X.500 address is: `n" 
    Write-Host -ForegroundColor Green $X500 `n
    Write-Host -ForegroundColor DarkCyan "Here is the Set-Mailbox command to add the X.500 address to a user (change the Identity attribute accordingly): `n"
    Write-Host -ForegroundColor Green "Set-Mailbox -Identity first.last@domain.com -EmailAddresses @{add=`"$X500`"}" `n
    Write-Host -ForegroundColor Yellow "Done!`n"
}