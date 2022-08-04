###################################################################################################################################################
###                                                                                                                                             ###
###  	Script by Corey St. Pierre, Ahead, LLC                                                                                                  ###
###                                                                                                                                             ###
###     Webpage -                      https://www.linkedin.com/in/coreystpierrebr/                                                             ###
###     GitHub Scripts -               https://github.com/roundtower-euc-microsoft/PowerShell_Scripts/tree/master/Production                    ###
###                                                                                                                                             ###
###                                                                                                                                             ###
###     Version -                      Version 1.0                                                                                              ###
###     Version History                Version 1.0 - 5/11/2020                                                                                  ###
###                                    Version 1.1 - Added PowerShell variable - to prevent truncation of results                               ###
###                                                                                                                                             ###
###                                                                                                                                             ###
###                                                                                                                                             ###
###                                                                                                                                             ###
###################################################################################################################################################

##############################################################################################################################
###                                                                                                                        ###
###  	Script Notes                                                                                                       ###
###     Script has been created to document the current local Exchange environment                                         ###
###     Script has been tested on Exchange 2010,2013,2016,2019                                                             ###
###                                                                                                                        ###
###     *** Important - Run this script in Exchange Management Shell                                                       ###
###                                                                                                                        ###
###     Update the variable - $logpath - to set the location you want the reports to be generated                          ###
###                                                                                                                        ###   
###     The script generates a separate report for each of the following for all the local Exchange servers                ###
###     in your Organization:                                                                                              ###
###                                                                                                                        ###
###      Exchange SSL certificates                                                                                         ###
###      OWA Virtual Directory URL                                                                                         ###
###      ActiveSync Virtual Directory URL                                                                                  ###
###      Outlook Anywhere configuration                                                                                    ###
###      AutoDiscover Virtual Directory URL                                                                                ###
###      OAB Virtual Directory URL                                                                                         ###
###      Web Services Virtual Directory URL                                                                                ###
###      Accepted Domains                                                                                                  ###
###      Email Address Policy configuration                                                                                ###
###      Receive Connectors configuration                                                                                  ###
###      Send Connectors configuration                                                                                     ###
###      Transport configuration                                                                                           ###
###      Mailbox Database configuration                                                                                    ###
###      Exchange Server configuration, including Exchange version                                                         ###
###      OWA Mailbox Policies                                                                                              ###  
###      Mobile Device Policies                                                                                            ###
###      Transport Rules                                                                                                   ###
###      Exchange Administrators                                                                                           ###
###      Mailbox Details                                                                                                   ###
###      Mailboxes with Forwarders                                                                                         ###
###      Mailboxes with Full Access Delegates                                                                              ###
###      Mailboxes with Send As Delegates                                                                                  ###
###      Mailboxes with Send on Behalf Delegates                                                                           ###
###      Mailbox statistics                                                                                                ###
###      Distrtibution Groups                                                                                              ###
###                                                                                                                        ###
##############################################################################################################################


### Update the log path variables below before running the script ####
$TestPath = Test-Path C:\ExchangeReports
if ($TestPath -eq $false)
    {
        New-Item -Path "C:\" -Name "ExchangeReports" -ItemType "directory"
    }
$logpath = "c:\ExchangeReports"

########################################################

### Do not change the variables below

$Mailboxes = get-mailbox -ResultSize Unlimited

$FormatEnumerationLimit=-1

########################################################

Import-Module ActiveDirectory

$Mailboxes | Get-ADPermission | where {($_.ExtendedRights -like "*Send-As*") -and ($_.IsInherited -eq $false) -and -not ($_.User -like "NT AUTHORITY\SELF")} | Select Identity,User,RecipientTypeDetails | Export-Csv -NoTypeInformation "$logpath\MailboxSendAsAccess-LocalExchange.csv"

$Mailboxes | Where-Object {$_.GrantSendOnBehalfTo} | select Name,@{Name='GrantSendOnBehalfTo';Expression={($_ | Select -ExpandProperty GrantSendOnBehalfTo | Select -ExpandProperty Name) -join ","}} | export-csv -notypeinformation "$logpath\MailboxSendOnBehalf-LocalExchange.csv"

$Mailboxes | Get-MailboxPermission | Where { ($_.IsInherited -eq $False) -and -not ($_.User -like “NT AUTHORITY\SELF”) -and -not ($_.User -like '*Discovery Management*') } | Select Identity, user,RecipientTypeDetails | Export-Csv -NoTypeInformation "$logpath\MailboxFullAccess-LocalExchange.csv"

########################################################

Get-ExchangeCertificate | Where {($_.IsSelfSigned -eq $False)} | Select CertificateDomains, Issuer, NotAfter, RootCAType, Services, Status, Subject | Out-File "$logpath\ExchangeCertificate-LocalExchange.txt" -NoClobber -Append

Get-OwaVirtualDirectory | Select Name,Server,InternalURL,ExternalURL  | FL | Out-File "$logpath\OWA-VirtualDirectory-LocalExchange.txt"

Get-PowerShellVirtualDirectory | Select Name,Server,InternalURL,ExternalURL | FL | Out-File "$logpath\PowerShellVirtualDirectory-LocalExchange.txt"

Get-ActiveSyncVirtualDirectory | Select Name,Server,InternalURL,ExternalURL | FL | Out-File "$logpath\ActiveSyncVirtualDirectory-LocalExchange.txt"

Get-ClientAccessServer | Select  Name,AutoDiscoverServiceCN,AutoDiscoverServiceInternalUri,OutlookAnywhereEnabled | FL | Out-File "$logpath\AutoDiscoverSCPandOutlookAnywhere-LocalExchange.txt"

Get-OabVirtualDirectory | Select Name,Server,InternalURL,ExternalURL | FL | Out-File "$logpath\OABVirtualDirectory-LocalExchange.txt"

Get-WebServicesVirtualDirectory | Select Name,Server,InternalURL,ExternalURL | FL | Out-File "$logpath\WebServicesVirtualDirectory-LocalExchange.txt"

Get-AcceptedDomain | Select Name,DomainName,DomainType,Default | Out-File "$logpath\AcceptedDomains-LocalExchange.txt"

Get-EmailAddressPolicy | Select Name,Priority,RecipientFilter,RecipientFilterApplied,IncludeRecipients,EnabledPrimarySMTPAddressTemplate,EnabledEmailAddressTemplates,Enabled,IsValid | Out-File "$logpath\EmailAddressPolicy-LocalExchange.txt"

Get-ReceiveConnector | Select Name,Enabled,ProtocolLoggingLevel,FQDN,MaxMessageSize,Bindings,RemoteIPRanges,AuthMechanism,PermissionGroups | Out-File "$logpath\ReceiveConnectors-LocalExchange.txt"

Get-SendConnector | Select Name,Enabled,ProtocolLoggingLevel,SmartHostsString,FQDN,MaxMessageSize,AddressSpaces,SourceTransportServers |  Out-File "$logpath\SendConnectors-LocalExchange.txt"

Get-TransportService | Select Name,InternalDNSServers,ExternalDNSServers,OutboundConnectionFailureRetryInterval,TransientFailureRetryInterval,TransientFailureRetryCount,MessageExpirationTimeout,DelayNotificationTimeout,MaxOutboundConnections,MaxPerDomainOutboundConnections,MessageTrackingLogEnabled,MessageTrackingLogPath,ConnectivityLogEnabled,ConnectivityLogPath,SendProtocolLogPath,ReceiveProtocolLogPath | Out-File "$logpath\TransportConfiguration-LocalExchange.txt"

Get-Mailboxdatabase | Select Servers,Name,EDBFilePath,LogFolderPath,MaintenanceSchedule,JournalRecipient,CircularLoggingEnabled,IssueWarningQuota,ProhibitSendQuota,ProhibitSendReceiveQuota,DeletedItemRetention,MailboxRetention,RetainDeletedItemsUntilBackup,OfflineAddressBook,LastFullBackup,LastIncrementalBackup,LastDifferentialBackup,DatabaseSize | Out-File "$logpath\MailboxDatabaseConfigs-LocalExchange.txt"

Get-ExchangeServer | Select Name,Server,Domain,FQDN,ServerRole,IsMemberOfCluster,AdminDisplayVersion | Out-File "$logpath\ExchangeServer-LocalExchange.txt"

Get-OwaMailboxPolicy | Select Name,ActiveSyncIntegrationEnabled,AllAddressListsEnabled,CalendarEnabled,ContactsEnabled,JournalEnabled,JunkEmailEnabled,RemindersAndNotificationsEnabled,NotesEnabled,PremiumClientEnabled,SearchFoldersEnabled,SignaturesEnabled,SpellCheckerEnabled,TasksEnabled,ThemeSelectionEnabled,UMIntegrationEnabled,ChangePasswordEnabled,RulesEnabled,PublicFoldersEnabled,SMimeEnabled,RecoverDeletedItemsEnabled,InstantMessagingEnabled,TextMessagingEnabled,DirectFileAccessOnPublicComputersEnabled,WebReadyDocumentViewingOnPublicComputersEnabled,DirectFileAccessOnPrivateComputersEnabled,WebReadyDocumentViewingOnPrivateComputersEnabled | Out-File "$logpath\OWAMailboxPolicies-LocalExchange.txt"

Get-MobileDeviceMailboxPolicy | Select Name,AllowNonProvisionableDevices,DevicePolicyRefreshInterval,PasswordEnabled,MaxCalendarAgeFilter,MaxEmailAgeFilter,MaxAttachmentSize,RequireManualSyncWhenRoaming,AllowHTMLEmail,AttachmentsEnabled,AllowStorageCard,AllowCameraTrue,AllowWiFi,AllowIrDA,AllowInternetSharing,AllowRemoteDesktop,AllowDesktopSync,AllowBluetooth,AllowBrowser,AllowConsumerEmail,AllowUnsignedApplications,AllowUnsignedInstallationPackages | Out-File "$logpath\MobileDevicePolicies-LocalExchange.txt"

Get-TransportRule | Select Name,Priority,Description,Comments,State | Out-File "$logpath\TransportRules-LocalExchange.txt"

Get-RoleGroupMember "Organization Management" | Out-File "$logpath\ExchangeAdmins-LocalExchange.txt"

### The following scripts output mailbox statistics ###

$MailboxStats = $Mailboxes | group-object recipienttypedetails | select count, name
$MailboxStats | Out-File "$logpath\MailboxStats-LocalExchange.txt"

### The following scripts output mailbox details including database ###
$Mailboxes | Select DisplayName,Alias,PrimarySMTPAddress,Database | export-csv -NoTypeInformation "$logpath\MailboxDetails-LocalExchange.csv"

### The following scripts output any forwarders configured on mailboxes ###
$Mailboxes | Where {($_.ForwardingAddress -ne $Null) -or ($_.ForwardingsmtpAddress -ne $Null)} | Select Name, DisplayName, PrimarySMTPAddress, UserPrincipalName, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward | export-csv -NoTypeInformation "$logpath\MailboxesWithForwarding-LocalExchange.csv"

### Get all groups into temp variable ###
$groups = Get-DistributionGroup -ResultSize Unlimited

### Export 1 export all distribution groups and a few settings ###
$groups | Select-Object RecipientTypeDetails,Name,Alias,DisplayName,PrimarySmtpAddress,@{name="SMTP Domain";expression={$_.PrimarySmtpAddress.Domain}},MemberJoinRestriction,MemberDepartRestriction,RequireSenderAuthenticationEnabled,@{label="ManagedBy";expression={[string]($_.managedby | foreach {$_.tostring().split("/")[-1]})}},@{name=”AcceptMessagesOnlyFrom”;expression={$_.AcceptMessagesOnlyFrom.Name -join “;”}},@{name=”AcceptMessagesOnlyFromDLMembers”;expression={$_.AcceptMessagesOnlyFromDLMembers.Name -join “;”}},@{name=”AcceptMessagesOnlyFromSendersOrMembers”;expression={$_.AcceptMessagesOnlyFromSendersOrMembers.Name -join “;”}},@{name=”ModeratedBy”;expression={$_.ModeratedBy.Name -join “;”}},@{name=”BypassModerationFromSendersOrMembers”;expression={$_.BypassModerationFromSendersOrMembers.Name -join “;”}},@{Name="GrantSendOnBehalfTo";Expression={$_.GrantSendOnBehalfTo.Name -join “;”}},ModerationEnabled,SendModerationNotifications,@{Name="EmailAddresses";Expression={$_.EmailAddresses -join “;”}} | Export-Csv $logpath\distributiongroups.csv -NoTypeInformation

Write-Host -ForegroundColor Green -BackgroundColor Black "Exchange Information has been gathered, please send this back to Ahead for processing"
Start-Sleep 3
ii C:\ExchangeReports