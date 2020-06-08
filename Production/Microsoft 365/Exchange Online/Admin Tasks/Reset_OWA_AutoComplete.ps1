<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.
#>

#requires -Version 2

<#
 	.SYNOPSIS
        This script is used to remove deleted mailboxes' address from all mailboxes' autocomplete list in OWA. 
    .DESCRIPTION
        This script is used to remove deleted mailboxes' address from all mailboxes' autocomplete list in OWA.
    .PARAMETER  Credential
        Indicates the credential to use for EWS service and connecting to PowerShell of Exchange Online. 
    .PARAMETER  RemovedSMTPAddress
        Indicate the un-existed mailbox address
    .EXAMPLE
        RemoveUnexistedAddressFromAutocomplete.ps1 -RemovedSMTPAddress removeduser@contoso.com
        The address removeduser@contoso.com is no longer existed in organization. This script will remove this address from OWA's autocomplete in all mailboxes. 
#>
Param
(
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.PSCredential]$Credential,
    [Parameter(Mandatory = $true)]
    [string] $RemovedSMTPAddress
)

Begin
{
    $webSvcInstallDirRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Exchange\Web Services\2.0" -PSProperty "Install Directory" -ErrorAction:SilentlyContinue
    if ($webSvcInstallDirRegKey -ne $null) {
	    $moduleFilePath = $webSvcInstallDirRegKey.'Install Directory' + 'Microsoft.Exchange.WebServices.dll'
	    Import-Module $moduleFilePath
    } 
    else 
    {
	    $errorMsg = "Please install Exchange Web Service Managed API 2.0"
	    throw $errorMsg
        Exit
    }

    $existingExSvcVar = (New-Variable -Name exService -Scope Global -ErrorAction:SilentlyContinue) -ne $null
		
	#Establish the connection to Exchange Web Service
	if ((-not $existingExSvcVar) -or $Force) 
    {
		$verboseMsg = $Messages.EstablishConnection
		$PSCmdlet.WriteVerbose($verboseMsg)
        if ($tzInfo -ne $null) {
            $exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
				    		[Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion,$tzInfo)			
        } else {
            $exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
				    		[Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion)
        }
			
		#Set network credential
		$userName = $Credential.UserName
		$exService.Credentials = $Credential.GetNetworkCredential()
            
        #Set exEWSUsername as global variable, in order to check RBAC permission of this account
        New-Variable -Name exEWSUsername -Scope Global -ErrorAction:SilentlyContinue
        Set-Variable -Name exEWSUsername -Value $Credential.UserName -Scope Global -Force
		Try
		{
			#Set the URL by using Autodiscover
			$exService.AutodiscoverUrl($userName,{$true})
			$verboseMsg = $Messages.SaveExWebSvcVariable
			$PSCmdlet.WriteVerbose($verboseMsg)
			Set-Variable -Name exService -Value $exService -Scope Global -Force
		}
		Catch [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverRemoteException]
		{
			$PSCmdlet.ThrowTerminatingError($_)
		}
		Catch
		{
			$PSCmdlet.ThrowTerminatingError($_)
		}
	} 
    else 
    {
		$verboseMsg = "Found Exchange Web Service variable in gloabl scope."
        $verboseMsg = $verboseMsg -f $exService.Credentials.Credentials.UserName
		$PSCmdlet.WriteVerbose($verboseMsg)            
	}
    
    $existingSession = Get-PSSession -Verbose:$false | Where-Object {$_.ConfigurationName -eq "Microsoft.Exchange"}
    if ($existingSession -eq $null) 
    {
        $verboseMsg = "Creating a new session to https://ps.outlook.com/powershell."
        $pscmdlet.WriteVerbose($verboseMsg)
        $O365Session = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri "https://ps.outlook.com/powershell" -Credential $Credential `
        -Authentication Basic -AllowRedirection
        #If session is newly created, import the session.
        Import-PSSession -Session $O365Session -Verbose:$false
        $existingSession = $O365Session
    } 
    else 
    {
        $verboseMsg = "Found existing session, new session creation is skipped."
        $pscmdlet.WriteVerbose($verboseMsg)
    }
}

Process
{
    $MailboxList = Get-Mailbox
    $checkRolePer = Get-ManagementRoleAssignment -RoleAssignee $exEWSUsername
    foreach ($roleItem in $checkRolePer)
    {
        if(($roleItem.Role -eq "ApplicationImpersonation") -and ($roleItem.RoleAssignmentDelegationType -eq "Regular"))
    {
        $hasPermission = $true
        break
    }
    }
    if($hasPermission -ne $true)
    {
        
        Write-Error "This account doesn't have the ApplicationImpersonation permission,refer http://msdn.microsoft.com/en-us/library/office/bb204095(v=exchg.140).aspx"
    }


    foreach($mailbox in $MailboxList)
    {
        $Mailbox.PrimarySmtpAddress
        $exService.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $Mailbox.PrimarySmtpAddress)
        
        #Remove from Recipient Cache
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecipientCache,$Mailbox.PrimarySmtpAddress)   
        $RecipientCache = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exService,$folderid)

        $psPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
        #Define ItemView to retrive just 1000 Items    
        $ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)    
        $fiItems = $null
        $SearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ContactSchema]::EmailAddress1, $RemovedSMTPAddress)
        if($exService.FindItems($RecipientCache.Id,$SearchFilter,$ivItemView) -ne $null)
        {
            $fiItems = $exService.FindItems($RecipientCache.Id,$SearchFilter,$ivItemView)
        }
        foreach($Item in $fiItems.Items)
        {
            $Item.delete("HardDelete")
        }

        #Remove from Suggested Contacts
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$Mailbox.PrimarySmtpAddress) 
        $fvFolderView = new-object Microsoft.Exchange.WebServices.Data.FolderView(1)
        $SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"Suggested Contacts")
        $findFolderResults = $exService.FindFolders($folderid,$SfSearchFilter,$fvFolderView)
        if($findFolderResults.Folders.Count -gt 0)
        {
            $SearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.ContactSchema]::EmailAddress1, $RemovedSMTPAddress)
            if($exService.FindItems($findFolderResults.Folders[0].Id,$SearchFilter,$ivItemView) -ne $null)
            {
                $fiItems = $exService.FindItems($findFolderResults.Folders[0].Id,$SearchFilter,$ivItemView)
            }
            $fiItems.count
            foreach($Item in $fiItems.Items)
            {
                $Item.delete("HardDelete")
            }
        }

        #Remove from OWA autocomplete cache
        $folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$Mailbox.PrimarySmtpAddress)    
        Try
        { 
            $UsrConfig = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind($exService, "OWA.AutocompleteCache", $folderid, [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)
        }
        Catch
        {}
        if($Error[0].Exception -eq "The specified object was not found in the store.")
        {
            $StringData = [System.Text.Encoding]::UTF8.GetString($UsrConfig.XmlData).Substring(1)
            $xmlDoc = New-Object System.Xml.XmlDocument
            if($StringData -ne "")
            {
                $XmlDoc.LoadXml($StringData)
                $nodes = $xmlDoc.SelectNodes("/AutoCompleteCache/entry")
                foreach($node in $nodes)
                {
                    if($node.smtpAddr -ne $null -and $node.smtpAddr.Contains($RemovedSMTPAddress))
                    {
                        Write-Host "contains"
                    }
                }
         
                #Convert the xml back into a byte array
                $UpdatedData = [System.Text.Encoding]::UTF8.GetBytes([System.Text.Encoding]::UTF8.GetString($UsrConfig.XmlData).Substring(0,1) + $XmlDoc.OuterXml)
                $UsrConfig.XmlData = $UpdatedData
 
                #Save the config back to the users mailbox
                $UsrConfig.Update()  
            }
        }
        $StringData = ""      
    }

}

End
{
}