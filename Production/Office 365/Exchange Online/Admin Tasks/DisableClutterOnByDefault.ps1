<#
        .SYNOPSIS
        A controller script that will set Clutter to manually disabled for all mailboxes
    
        .DESCRIPTION
        Beginning in June, 2015 all Office 365 mailboxes will have the new Clutter feature enabled by default.
        Administrators have the ability to manually alter this behavior so that a user must explicitly enable
        Clutter. This script will disable the 'On by default' nature of Clutter for all existing mailboxes
        that do not already have Clutter enabled. The script also outputs some information for each mailbox:

        PrimarySMTPAddress: the email address for the current mailbox
        ClutterEnabled: whether Clutter has been manually enabled
        Action: What action the script has taken. Values are 'None','Disable', and 'FAIL'.

        This information can easily be exported to CSV or XML for logging of status and actions.
    
        A failed action indicates that there was a problem retrieving or setting the status of Clutter. This
        usually indicates that the server your Powershell remote session is executing against or the server the 
        mailbox resides on is on a previous build that does not support the Clutter cmdlets.

        .EXAMPLE
        $Credential = Get-Credential
        & .\DisableClutterOnByDefault.ps1 -Credential $Credential -Whatif

        The above commands will execute the script without making any changes - you'll just get an idea what
        actions will be taken and for whom. Always do this first.

        .EXAMPLE
        & .\DisableClutterOnByDefault.ps1 -Credential $Credential

        This command will execute the script and output the resulting actions to the console.

        .EXAMPLE
        & .\DisableClutterOnByDefault.ps1 -Credential $Credential | Export-CSV .\Clutter.csv -NoTypeInformation

        This command will execute the script and output the resulting actions to a CSV file for review.

        .LINK
        http://psescape.azurewebsites.net/exchange-online-managing-clutter

        .NOTES
        Author: Matt McNabb
        Date: 5/25/2015
        DISCLAIMER: This script is provided 'AS IS'. It has been tested for personal use, please   
        test in a lab environment before using in a production environment.

#>

[CmdletBinding(SupportsShouldProcess = $true)]
[OutputType([PSObject])]
param
(
    # A credential for an Office 365 Exchange Online Service Administrator
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.CredentialAttribute()]
    $Credential
)

# Connect to Exchange Online
$Splat = @{
    ConfigurationName = 'Microsoft.Exchange'
    ConnectionUri     = 'https://outlook.office365.com/powershell-liveid/'
    Credential        = $Credential
    Authentication    = 'Basic'
    AllowRedirection  = $true
}
$Session = New-PSSession @Splat
$null = Import-PSSession $Session -WarningAction SilentlyContinue

# Check that the Clutter cmdlets exist in the Exchange Online module
try { $null = Get-Command -Name Get-Clutter -ErrorAction Stop }
catch [System.Management.Automation.CommandNotFoundException]
{
    throw "Clutter cmdlets not found! Microsoft is trying hard to remedy this - please try again later."
}

# Implicit remoting doesn't honor the -ErrorAction parameter, so we have to set the global
# preference to 'Stop' temporarily
$EAPPrevious = $ErrorActionPreference
$Global:ErrorActionPreference = 'Stop'

# Find mailboxes and disable Clutter
$Mailboxes = Get-Mailbox -Filter * -ResultSize Unlimited

foreach ($Mailbox in $Mailboxes)
{
    $PrimarySMTPAddress = $Mailbox.PrimarySMTPAddress
    $Hash = [ordered]@{
        PrimarySMTPAddress  = $PrimarySMTPAddress
        ClutterEnabled      = $null
        Action              = $null
    }
    try
    {
        $ClutterEnabled = Get-Clutter -Identity $PrimarySMTPAddress | Select-Object -ExpandProperty isEnabled
        $Hash.ClutterEnabled = $ClutterEnabled
        
        if ($PSCmdlet.ShouldProcess($Hash.PrimarySMTPAddress))
        {
            # Make sure Clutter wasn't enabled by the user
            if (!$Hash.ClutterEnabled)
            {
                $null = Set-Clutter -Identity $PrimarySMTPAddress -Enable $false -ErrorAction Stop
                $Hash.Action = 'Disable'
            }
        }
    }
    catch [System.InvalidOperationException]
    {
        # Get-Clutter failed - command does not exist on the server
        $Hash.ClutterEnabled = 'FAIL'
        $Hash.Action = $null
    }
    catch [System.Management.Automation.RemoteException]
    {
        # Set-Clutter failed
        $Hash.Action = 'FAIL'
    }
    
    New-Object -TypeName PSObject -Property $Hash
}

$Global:ErrorActionPreference = $EAPPrevious

Remove-PSSession $Session
Remove-Module tmp*

#requires -Version 2.0