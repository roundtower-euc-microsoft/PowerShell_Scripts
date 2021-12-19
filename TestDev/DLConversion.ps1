#requires -version 2
<#

.DISCLAIMER

    PLEASE USE WITH CAUTION!!!! SCRIPTS ARE NOT SUPPORTED BY MICROSOFT IN ANY WAY, SHAPE,
    OR FORM AS PER THEIR TERMS AND CONDITIONS. AHEAD, LLC AND THE SCRIPTS AUTHORS WILL
    PROVIDE PERIODIC UPDATES AND MAINTENANCE TO THIS SCRIPT, HOWEVER, ARE NOT LIABLE FOR
    UNAUTHORIZED CHANGES NOT DOCUMENTED IN THE SCRIPT OR ANY IMPROPER USAGE OF SAID SCRIPTS.


.SYNOPSIS

This script is designed to assist users in migrating distributon lists to Office 365.
	  
************************************************************************************************
************************************************************************************************
************************************************************************************************
The PSLOGGING module must be installed / updated prior to use.
---install-module PSLOGGING
---update-module PSLOGGING.
************************************************************************************************
************************************************************************************************
************************************************************************************************

Information regarding module can be found here:

https://www.powershellgallery.com/packages/PSLogging/2.5.2
https://9to5it.com/powershell-logging-v2-easily-create-log-files/

All credits for the logging module to the owner and many thanks for allowing it's public use.

.DESCRIPTION

The script performs the following functions:

1)	Log into remote powershell to On-Premises Exchange.
2)	Log into remote powershell to Office 365 / Exchange Online and prefix the command with o365.

Note:	This may require enabling basic authentication on one or more endpoints for powershell remoting.

3)	Obtain the distribution list name from the command line.
4)	Validate the DL exists on premsies.	[Hard fail if not found]
5)	Validate the DL exists in Office 365. [Hard fail if not found]
6)	Capture the distribution list membership on premises and validate that each recipient exists in Exchange Online [Hard fail if recipient not found]
7)	Capture all multivalued attributs of the distribution list, convert the list to primary SMTP addresses, and validate those recipients exist in Exchange Online [Hard fail if recipient not found]
8)	Move the distribution list to a non-sync OU.
9)	Trigger remotely a delta sync of AD connect to remove the distribution list from Exchange Online.
10)	Validate the DL was removed from Exchange Online.
11)	Recreate the distribution list directly in Exchange Online.
12)	Stamp all attributes of the on-premises DL to the new DL in Exchange Online.
13)	Stamp the original legacyExchangeDN to the new DL as an X500 address to preserve reply to functionality.
14) Assuming convert to contact logic is not overridden continue...
15)  Delete the DL moved to the converted DL OU.
16)  Create a dynamic distribution group where the criteria matches a contact that will be created.
17)  Create a mail contact that uses the onmicrosoft.com address of the group and has custom attributes that match the dynamic DL created.

.INPUTS

Credential files are required to be created prior to script execution as secure string XML files.
Administrator is responsible adjusting the following variables to meet their environment.

Variables proceeded by ##ADMIN##

.OUTPUTS

Operations log file.
XML files for backup of on premises and cloud DLs.

.NOTES

Version:        1.0
Author:         Corey St. Pierre, Ahead, LLC

Script requires PowerShell version 2 minimum to run

.CHANGE LOG
Creation Date:  December 10,2021
Purpose/Change: Initial script development

Version:        1.1
Author:         Corey St. Pierre
Purpose/Change: Updated group creation function to allow user override through parameter.  This allows the user to override if the group is provisioned as security or distribution regardless of on-premises representation.

Version:        1.2
Author:         Corey St. Pierre
Purpose/Change: Implemented new parameter to allow the group to be converted to a mail enabled contact.  During testing of these changes discovered that a common scenario was not accounted for.  It is possible on the multi-valued attributes that they coudl be set to the same group being migrated.   For example group MIGRATETEST only accepts messages from MIGRATETEST.
In our testing code we tested to see if any groups on these attributes were already migrated - and of course the group we're migrating has not yet been migrated.

Version:		1.3
Author:			Corey St. Pierre
Purpose/Change:	Implemented some code changes to address some issues.  The first issue that was discovered came with the choice to convert the on premises DL to a mail enabled contacts.  If the group on prmeises was a security group...
remove-distributionGroup is unable to remove the group unless the person executing the script was also a manager of the group.  The code was changed to implement the -bypassSecurityGroupManagerCheck which allows the executor to remove the DL.
The script also when provisioning the mail enabled contact used to mirror as many of the DL attributes as possible.  For example, accept messages from / reject messages from etc.  When these attribute were mirrored this was fine...
Until it was realized that they could adjust in the service and there was no way to mirror them back to the contact.  It made more sense to create the contact with the simple attributes - and let the message move onto the service - 
where further decisions to implement accept / reject / grant etc could be evaluated and would then be managed and up to date.  The last change taken was to adjust the timing of creating the mail enabled contact.
In testing with a multi-DC environment we discovered that between deleting the DL and creating the contact sometimes the AD cache was not updated - and reuslted in an error provisining the contact that SMTP addresses exist.
Reused the one minute timout logic + dc replication between deleting the DL <and> provisioning the contact which corrected this issue.

Version:		1.4
Author:			Corey St. Pierre
Purpose/Change:	In receiving customer feedback an interesting scenario was presented.  In versions prior to 1.4 convert to contact was not mandatory as true.  This would leave the distribution group on premises.  In recent weeks a customer performing conversions accidentally rebuilt their ad connect box, and selected the OU contacining converted groups.
In doing so - the groups softmatched from on prmeises to the migrated cloud only distribution lists.  This resulted in the migration being undone - and old information from on premises now overwriting new information in the cloud.  The first change in this build
was to make the convert to contact TRUE.  Customers desiring to retain the group on premises must now specify -convertToContact:$FALSE.

In the build array logic the powershell session was rebuilt for each iteration.  This added time and was not necessary if the individual array items were less that 1000 objects.  We now only refresh if the array type is greater than 1000 objects.

Created a remote powershell session to the domain controller specified / now required.  This is where we will perform our work.

Updated the move to ou to function to utilize the powershell session and static domain controller specified.

In the ad connect invocation function split the importation of the module from the delta sync so that individual errors could be trapped.

Redefined the replicate domain controllers function.  It's now broken down into triggering inbound replication and outbound replication from local domain controllers within the site.

Version:		1.5
Author:			Corey St. Pierre
Purpose/Change:	In this version we correct some of operations.  For example, we update the AD calls to utilize calls that work assuming a non-alias is utilized for the group.
Additionally new functions have been created to log information regaridng the items created.

Version:		1.6
Author:			Corey St. Pierre
Purpose/Change:	In this version we are implementing a switch that allows administrators to bypass retaining on premsies settings for the group.
When this switch is utilized as FALSE - if the migrated group was set on any other disrtibution list on premises with permissions or is the member of another distribution group
these settings are lost.  This increases script execution time - but also results in potentially loosing effective permissions or group memberships on premises.
Recommendations to only utilize this switch when a simple distribution group is being migrated.

Version:		1.7
Author:			Corey St. Pierre
Purpose/Change:	In this version we implement a swtich allowing for credentials to be supplied in an interactive fashion.  This prevents having to supply credentials within the XML file.

Version:		1.8
Author:			Corey St. Pierre
Purpose/Change: In version 1.8 we will attempt to track cloud only settings as they apply to the distribution group.  For example, it is possible that the on premsies distribution group was added
with permissions, membership, restrictions, forwarding to cloud only objects or cloud only settings (like forwarding).  As with tracking on premises changes - the size of the tenant
directly impacts the ability to do this work in a timely fashion.  A switch is utilized to override tracking these should the administrator not care.

Version:		1.8.1
Author:			Corey St. Pierre
Purpose/Chnage: Correcting get-distributionGroupMember for multiple domains.  In the current confige using -domainController only returns members from the domain of the controller specified.
We need to ignore the scope and allow it to search all domains that Exchange is aware of.

Version:		1.8.2
Author:			Corey St. Pierre
Purpose/Change: Correcting for distribution lists that contain apostrophies.

Version:		1.8.3
Author:			Corey St. Pierre
Purpose/Change: Implementing changes to ad server settings to ensure view entire forest is true.
 
.EXAMPLE

DLConversion -dlToConvert dl@domain.com -ignoreInvalidDLMember:$TRUE -ignoreInvalidManagedByMember:$TRUE -groupTypeOverride "Security" -convertToContact:$TRUE -retainOnPremisesSettings:$TRUE -requireInteractiveCredentials:$TRUE
#>

#---------------------------------------------------------[Script Parameters]------------------------------------------------------

Param (
    #Script parameters go here

	[Parameter(Mandatory=$True,Position=1)]
    [string]$dlToConvert,
    [Parameter(Position=2)]
    [boolean]$ignoreInvalidDLMember=$FALSE,
    [Parameter(Position=3)]
	[boolean]$ignoreInvalidManagedByMember=$FALSE,
	[Parameter(Position=4)]
	[ValidateSet("Security","Distribution",$NULL)]
	[string]$groupTypeOverride=$NULL,
	[Parameter(Position=5)]
	[boolean]$convertToContact=$TRUE,
	[Parameter(Position=6)]
	[boolean]$retainOnPremisesSettings=$True,
	[Parameter(Position=7)]
	[boolean]$retainO365CloudOnlySettings=$True,
	[Parameter(Position=8)]
	[boolean]$requireInteractiveCredentials=$FALSE
)

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

#Set Error Action to Silently Continue
$ErrorActionPreference = "SilentlyContinue"

Import-Module PSLogging

#----------------------------------------------------------[Declarations]----------------------------------------------------------

#Script Version
$sScriptVersion = "1.8.2"

#Log File Info
<###ADMIN###>$script:sLogPath = "C:\Scripts\Working\"
<###ADMIN###>$script:sLogName = "DLConversion.log"
<###ADMIN###>$script:sLogFile = Join-Path -Path $sLogPath -ChildPath $sLogName

#Establish credential information.

<###ADMIN###>$script:credentialFilePath = "C:\Scripts\"  #Path to the local credential files.
<###ADMIN###>$script:onPremisesCredentialFileName = "OnPremises-Credentials.cred"  #OnPremises credential XML file.
<###ADMIN###>$script:office365CredentialFileName = "Office365-Credentials.cred"  #Office365 credential XML file.
$script:onPremisesCredentialFile = Join-Path -Path $script:credentialFilePath -ChildPath $script:onPremisesCredentialFileName  #Full path and file name to onpremises credential file.
$script:office365CredentialFile = Join-Path -Path $script:credentialFilePath -ChildPath $script:office365CredentialFileName  #Full path and file name to Office 365 credential file.
$script:onPremisesCredential = $NULL  #Onpremises credentials
$script:office365Credential = $NULL  #Office 365 creentials

#Establish powershell server references.

<###ADMIN###>$script:onPremisesAADConnectServer = "server.domain.local"  #FQDN of the local AD Connect instance.
<###ADMIN###>$script:onPremisesPowerShell = "https://server.domain.com/powershell"  #URL of on premises powershell instance.
$script:office365Powershell = "https://outlook.office365.com/powershell-liveID/"  #URL of Office 365 powershell instance.

#Establish global variables for powershell sessions.

$script:office365PowershellConfiguration =  "Microsoft.Exchange" #Office 365 powershell configuration.
$script:onPremisesPowershellConfiguration =  "Microsoft.Exchange" #Exchange ON premises powershell configuration.
$script:office365PowershellAuthentication = "Basic" #Office 365 powershell authentication.
$script:onPremisesPowershellAuthentication = "Kerberos" #Exchange on premises powershell authentication.
$script:office365PowerShellSession = $null #Office 365 powershell session.
$script:onPremisesPowerShellSession = $null #Exchange on premises powershell session.
$script:onPremisesAADConnectPowerShellSession = $null #On premises aad connect powershell session.
$script:onPremisesADDomainControllerPowerShellSession = $null #On premises aad connect powershell session.

#Establish script variables for distribution list operations.

$script:onpremisesdlConfiguration = $NULL  #Gathers the on premises DL information to a variable.
$script:office365DLConfiguration = $NULL  #Gather the Office 365 DL information to a variable.
$script:onpremisesdlconfigurationMembership = $null #On premise dl membership.
$script:newOffice365DLConfiguration = $NULL
$script:newOffice365DLConfigurationMembership = $NULL
[array]$script:onpremisesdlconfigurationMembershipArray = @() #Array of psObjects that represent DL membership.
[array]$script:onpremisesdlconfigurationManagedByArray = @() #Array of psObjects that represent managed by membership.
[array]$script:onpremisesdlconfigurationModeratedByArray = @() #Array of psObjects that represent moderated by membership.
[array]$script:onpremisesdlconfigurationGrantSendOnBehalfTOArray = @() #Array of psObjects that represent grant send on behalf to membership.
[array]$script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers = @() #Array of psObjects that represent accept messages only from senders or members membership.
[array]$script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers = @() #Array of psObjects that represent reject messages only from senders or members membership.
[array]$script:onPremsiesDLBypassModerationFromSendersOrMembers = @() #Array of psObjects that represent bypass moderation only from senders or members membership.

$script:newOffice365DLConfiguration=$NULL
$script:x500Address=$NULL

#Establish script variables for active directory operations.

<###ADMIN###>$script:groupOrganizationalUnit = "OU=ConvertedDL,DC=DOMAIN,DC=LOCAL" #OU to move migrated DLs too.
<###ADMIN###>$script:adDomainController = "dcname.domain.com" #List of domain controllers in domain.
<###ADMIN###>[int32]$script:adDomainReplicationTime = 1 #Timeout to wait and allow for ad replication.
<###ADMIN###>[int32]$script:dlDeletionTime = 1 #Timeout to wait before rechecking for deleted DL.

#Establish script variables to backup distribution list information.

<###ADMIN###>$script:backupXMLPath = "C:\Scripts\Working\" #Location of backup XML files.
$script:archiveXMLPath = $NULL
<###ADMIN###>$script:onpremisesdlconfigurationXMLName = "onpremisesdlConfiguration.XML" #On premises XML file name.
<###ADMIN###>$script:office365DLXMLName = "office365DLConfiguration.XML" #Cloud XML file name.
<###ADMIN###>$script:onPremsiesDLConfigurationMembershipXMLName = "onpremisesDLConfigurationMembership.XML"
<###ADMIN###>$script:newOffice365DLConfigurationXMLName = "newOffice365DLConfiguration.XML"
<###ADMIN###>$script:newOffice365DLConfigurationMembershipXMLName = "newOffice365DLConfigurationMembership.XML"
<###ADMIN###>$script:onPremisesMemberOfXMLName = "onPremsiesMemberOf.XML"
<###ADMIN###>$script:originalGrantSendOnBehalfToXMLName="onPremsiesGrantSendOnBehalfTo.xml"
<###ADMIN###>$script:originalAcceptMessagesFromXMLName="onPremsiesAcceptMessagesFrom.xml"
<###ADMIN###>$script:originalManagedByXMLName="onPremsiesManagedBy.xml"
<###ADMIN###>$script:originalRejectMessagesFromXMLName="onPremsiesRejectMessagesFrom.xml"
<###ADMIN###>$script:originalBypassModerationFromSendersOrMembersXMLName="OnPremisesBypass.xml"
<###ADMIN###>$script:originalForwardingAddressXMLName="onPremisesForwardAddress.xml"
<###ADMIN###>$script:originalForwardingSMTPAddressXMLName="onPremisesForwardingSMTPAddress.xml"

<###ADMIN###>$script:onpremisesdlconfigurationMembershipArrayXMLName = "onpremisesdlconfigurationMembershipArray.xml"
<###ADMIN###>$script:onpremisesdlconfigurationManagedByArrayXMLName = "onpremisesdlconfigurationManagedBy.xml" 
<###ADMIN###>$script:onpremisesdlconfigurationModeratedByArrayXMLName = "onpremisesdlconfigurationModeratedBy.xml"
<###ADMIN###>$script:onpremisesdlconfigurationGrantSendOnBehalfTOArrayXMLName = "onpremisesdlconfigurationGrant.xml"
<###ADMIN###>$script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembersXMLName = "onpremisesdlconfigurationAccept.xml"
<###ADMIN###>$script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembersXMLName = "onpremisesdlconfigurationReject.xml" 
<###ADMIN###>$script:onPremsiesDLBypassModerationFromSendersOrMembersXMLName = "onPremsiesDLBypass.xml"

$script:onpremisesdlconfigurationMembershipArrayXMLPath = Join-Path $script:backupXMLPath -ChildPath $script:onpremisesdlconfigurationMembershipArrayXMLName
$script:onpremisesdlconfigurationManagedByArrayXMLPath = Join-Path $script:backupXMLPath -ChildPath $script:onpremisesdlconfigurationManagedByArrayXMLName 
$script:onpremisesdlconfigurationModeratedByArrayXMLPath = Join-Path $script:backupXMLPath -ChildPath $script:onpremisesdlconfigurationModeratedByArrayXMLName
$script:onpremisesdlconfigurationGrantSendOnBehalfTOArrayXMLPath = Join-Path $script:backupXMLPath -ChildPath $script:onpremisesdlconfigurationGrantSendOnBehalfTOArrayXMLName
$script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembersXMLPath = Join-Path $script:backupXMLPath -ChildPath $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembersXMLName
$script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembersXMLPath = Join-Path $script:backupXMLPath -ChildPath $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembersXMLName
$script:onPremsiesDLBypassModerationFromSendersOrMembersXMLPath = Join-Path $script:backupXMLPath -ChildPath $script:onPremsiesDLBypassModerationFromSendersOrMembersXMLName

$script:onPremisesXML = Join-Path $script:backupXMLPath -ChildPath $script:onpremisesdlconfigurationXMLName #Full path to on premises XML.
$script:office365XML = Join-Path $script:backupXMLPath -ChildPath $script:office365DLXMLName #Full path to cloud XML.
$script:onPremsiesMembershipXML = Join-Path $script:backupXMLPath -ChildPath $script:onPremsiesDLConfigurationMembershipXMLName
$script:newOffice365XML = Join-Path $script:backupXMLPath -ChildPath $script:newOffice365DLConfigurationXMLName
$script:newOffice365MembershipXML = Join-Path $script:backupXMLPath -ChildPath $script:newOffice365DLConfigurationMembershipXMLName
$script:onPremisesMemberOfXML = Join-Path $script:backupXMLPath -ChildPath $script:onPremisesMemberOfXMLName

$script:originalGrantSendOnBehalfToXML=Join-Path $script:backupXMLPath -ChildPath $script:originalGrantSendOnBehalfToXMLName
$script:originalAcceptMessagesFromXML=Join-Path $script:backupXMLPath -ChildPath $script:originalAcceptMessagesFromXMLName
$script:originalManagedByXML=Join-Path $script:backupXMLPath -ChildPath $script:originalManagedByXMLName
$script:originalRejectMessagesFromXML=Join-Path $script:backupXMLPath -ChildPath $script:originalRejectMessagesFromXMLName
$script:originalBypassModerationFromSendersOrMembersXML=Join-Path $script:backupXMLPath -ChildPath $script:originalBypassModerationFromSendersOrMembersXMLName
$script:originalForwardingSMTPAddressXML=Join-Path $script:backupXMLPath -ChildPath $script:originalForwardingSMTPAddressXMLName
$script:originalForwardingAddressXML=Join-Path $script:backupXMLPath -ChildPath $script:originalForwardingAddressXMLName

#Establish misc.

$script:aadconnectRetryRequired = $FALSE #Determines if ad connect sync retry is required.
$script:dlDeletionRetryRequired = $FALSE #Determines if deleted DL retry is required.
[int]$script:forCounter = $NULL #Counter utilized 
<###ADMIN###>[int32]$script:refreshCounter=1000

[array]$script:onPremisesDLMemberOf = @()	#Holds an array of groups that the migrated DL is a member of.
$script:onPremisesMovedDLConfiguration = $NULL	#Holds the seetings of the distribution group after the OU has changed.
[array]$script:originalGrantSendOnBehalfTo = @()  #Holds all distribution lists where the converted DL had grant send on behalf rights.
[array]$script:originalAcceptMessagesFrom = @()  #HOlds all the distribution lists where the converted DL had accept messages from rights.
[array]$script:originalRejectMessagesFrom = @()  #HOlds all the distribution lists where the converted DL had reject messages from rights.
[array]$script:originalForwardingAddress = @()
[array]$script:originalForwardingSMTPAddress = @()
[array]$script:originalBypassModerationFromSendersOrMembers = @()
[array]$script:originalManagedBy = @()
$script:randomContactName = $NULL
$script:remoteRoutingAddress = $NULL
$script:wellKnownSelfAccountSid = "S-1-5-10"
$script:onPremisesNewContactConfiguration = $NULL

$script:arrayCounter=0	#Counter used to build arrays for the recipient arrays.
$script:arrayGUID=$NULL	#Global used to store the recipient GUIDs to normalize objects.

$script:newDynamicDLAddress = $NULL	#Primary SMTP address built for the dynamic distribution list.

[array]$script:O365CloudOnlyDistributionGroups = $NULL	#Array holds all distribution groups that are cloud only distribution groups.
[array]$script:originalO365GrantSendOnBehalfTo = $NULL	#Array of all groups that are cloud only that contain the migrated DL as grant send on behalf to.
[array]$script:originalO365AcceptMessagesOnlyFromDLMembers = $NULL #Array of all groups that are cloud only that contain the mgirated DL as accept messages only from.
[array]$script:originalO365ManagedBy = $NULL #Array of all groups that are cloud only that contain the migrated DL as managed By.
[array]$script:originalO365RejectMessagesFromDLMembers = $NULL #Array of all groups that are cloud only that contain the migrated DL as reject messages from DL members.
[array]$script:originalO365BypassModerationFromSendersOrMembers = $NULL #Array of all groups that are cloud only that contain the migrated DL as bypass moderation.
[array]$script:originalO365ForwardingAddress = $NULL #Array of all objects where the migrated group is set for explicity forwarding address.
[array]$script:originalO365MemberOf = $NULL #Array of all cloud only distribution groups where the migrated group is a member of the distribution group.

<###ADMIN###>$script:originalO365GrantSendOnBehalfToXMLName="o365GrantSendOnBehalfTo.xml"
<###ADMIN###>$script:originalO365AcceptMessagesFromXMLName="o365AcceptMessagesFrom.xml"
<###ADMIN###>$script:originalO365ManagedByXMLName="o365ManagedBy.xml"
<###ADMIN###>$script:originalO365RejectMessagesFromXMLName="o365RejectMessagesFrom.xml"
<###ADMIN###>$script:originalO365BypassModerationFromSendersOrMembersXMLName="o365Bypass.xml"
<###ADMIN###>$script:originalForwardingAddressXMLName="o365ForwardAddress.xml"
<###ADMIN###>$script:originalO365MemberOfXMLName="o365MemberOf.xml"


$script:originalO365GrantSendOnBehalfToXML=Join-Path $script:backupXMLPath -ChildPath $script:originalo365GrantSendOnBehalfToXMLName
$script:originalO365AcceptMessagesFromXML=Join-Path $script:backupXMLPath -ChildPath $script:originalo365AcceptMessagesFromXMLName
$script:originalO365ManagedByXML=Join-Path $script:backupXMLPath -ChildPath $script:originalo365ManagedByXMLName
$script:originalO365RejectMessagesFromXML=Join-Path $script:backupXMLPath -ChildPath $script:originalo365RejectMessagesFromXMLName
$script:originalO365BypassModerationFromSendersOrMembersXML=Join-Path $script:backupXMLPath -ChildPath $script:originalo365BypassModerationFromSendersOrMembersXMLName
$script:originalO365ForwardingAddressXML=Join-Path $script:backupXMLPath -ChildPath $script:originalForwardingAddressXMLName
$script:originalO365MemberOfXML=Join-Path $script:backupXMLPath -ChildPath $script:originalO365MemberOfXMLName

[array]$script:originalO365GroupGrantSendOnBehalfTo = $NULL	#Array of all office 365 groups that are cloud only that contain the migrated DL as grant send on behalf to.
[array]$script:originalO365GroupAcceptMessagesOnlyFromDLMembers = $NULL #Array of all office 365 groups that are cloud only that contain the mgirated DL as accept messages only from.
[array]$script:originalO365GroupRejectMessagesFromDLMembers = $NULL #Array of all office 365 groups that are cloud only that contain the migrated DL as reject messages from DL members.

<###ADMIN###>$script:originalO365GroupGrantSendOnBehalfToXMLName="o365GroupGrantSendOnBehalfTo.xml"
<###ADMIN###>$script:originalO365GroupAcceptMessagesFromXMLName="o365GroupAcceptMessagesFrom.xml"
<###ADMIN###>$script:originalO365GroupRejectMessagesFromXMLName="o365GroupRejectMessagesFrom.xml"

$script:originalO365GroupGrantSendOnBehalfToXML=Join-Path $script:backupXMLPath -ChildPath $script:originalo365GroupGrantSendOnBehalfToXMLName
$script:originalO365GroupAcceptMessagesFromXML=Join-Path $script:backupXMLPath -ChildPath $script:originalo365GroupAcceptMessagesFromXMLName
$script:originalO365GroupRejectMessagesFromXML=Join-Path $script:backupXMLPath -ChildPath $script:originalo365GroupRejectMessagesFromXMLName

<###ADMIN###>[int]$script:refreshPowerShellSessionCounter=1000


#-----------------------------------------------------------[Functions]------------------------------------------------------------

<#
*******************************************************************************************************

Function ConvertFrom-DN 

.DESCRIPTION

This function takes a distinguished name and converts it to a cononical name.

All credits to the author (adapted by  Corey St. Pierre)

https://gist.github.com/joegasper/3fafa5750261d96d5e6edf112414ae18

.PARAMETER 

Distinguished Name

.INPUTS

NONE

.OUTPUTS 

Canonical Name

*******************************************************************************************************
#>

function ConvertFrom-DN 
{
    [cmdletbinding()]

    param(

    [Parameter(Mandatory,ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)] 

    [ValidateNotNullOrEmpty()]

    [string[]]$DistinguishedName

    )

    process 
    {
        foreach ($DN in $DistinguishedName) 
        {
            foreach ( $item in ($DN.replace('\,','~').split(","))) 
            {
                switch ($item.TrimStart().Substring(0,2)) 
                {
                    'CN' {$CN = '/' + $item.Replace("CN=","").Replace("/","\/")}

                    'OU' {$OU += ,$item.Replace("OU=","");$OU += '/'}

                    'DC' {$DC += $item.Replace("DC=","");$DC += '.'}

                }
            } 

            $CanonicalName = $DC.Substring(0,$DC.length - 1)

            for ($i = $OU.count;$i -ge 0;$i -- )
            {
                $CanonicalName += $OU[$i]
            }

            if ( $DN.Substring(0,2) -eq 'CN' ) 
            {
                $CanonicalName += $CN.Replace('~','\,')
			}
            Write-Output $CanonicalName
        }
    }
}

<#
*******************************************************************************************************

Function Start-PSCountdown

.DESCRIPTION

This function starts a visual countdown timer to show progress at scheduled waits.

All credits to the author

https://gist.github.com/jdhitsolutions/2e58d1aa41f684408b64488259bbeed0

.PARAMETER 

Multiple defined by author.

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>
function Start-PSCountdown {
    [cmdletbinding()]
    [OutputType("None")]
    Param(
        [Parameter(Position = 0, HelpMessage = "Enter the number of minutes to countdown (1-60). The default is 5.")]
        [ValidateRange(1, 60)]
        [int32]$Minutes = 5,
        [Parameter(HelpMessage = "Enter the text for the progress bar title.")]
        [ValidateNotNullorEmpty()]
        [string]$Title = "Counting Down ",
        [Parameter(Position = 1, HelpMessage = "Enter a primary message to display in the parent window.")]
        [ValidateNotNullorEmpty()]
        [string]$Message = "Starting soon.",
        [Parameter(HelpMessage = "Use this parameter to clear the screen prior to starting the countdown.")]
        [switch]$ClearHost 
    )
    DynamicParam {
        #this doesn't appear to work in PowerShell core on Linux
        if ($host.PrivateData.ProgressBackgroundColor -And ( $PSVersionTable.Platform -eq 'Win32NT' -OR $PSEdition -eq 'Desktop')) {
            #define a parameter attribute object
            $attributes = New-Object System.Management.Automation.ParameterAttribute
            $attributes.ValueFromPipelineByPropertyName = $False
            $attributes.Mandatory = $false
            $attributes.HelpMessage = @"
Select a progress bar style. This only applies when using the PowerShell console or ISE.           
Default - use the current value of `$host.PrivateData.ProgressBarBackgroundColor
Transparent - set the progress bar background color to the same as the console
Random - randomly cycle through a list of console colors
"@
            #define a collection for attributes
            $attributeCollection = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
            $attributeCollection.Add($attributes)
            #define the validate set attribute
            $validate = [System.Management.Automation.ValidateSetAttribute]::new("Default", "Random", "Transparent")
            $attributeCollection.Add($validate)
            #add an alias
            $alias = [System.Management.Automation.AliasAttribute]::new("style")
            $attributeCollection.Add($alias)
            #define the dynamic param
            $dynParam1 = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("ProgressStyle", [string], $attributeCollection)
            $dynParam1.Value = "Default"
            #create array of dynamic parameters
            $paramDictionary = New-Object -Type System.Management.Automation.RuntimeDefinedParameterDictionary
            $paramDictionary.Add("ProgressStyle", $dynParam1)
            #use the array
            return $paramDictionary     
        } #if
    } #dynamic parameter
    Begin {
        $loading = @(
            'Sleeping'
        )
		if ($ClearHost)
		{
            Clear-Host
        }
        $PSBoundParameters | out-string | Write-Verbose
        if ($psboundparameters.ContainsKey('progressStyle')) { 
            if ($PSBoundParameters.Item('ProgressStyle') -ne 'default') {
                $saved = $host.PrivateData.ProgressBackgroundColor 
            }
            if ($PSBoundParameters.Item('ProgressStyle') -eq 'transparent') {
                $host.PrivateData.progressBackgroundColor = $host.ui.RawUI.BackgroundColor
            }
        }
        $startTime = Get-Date
        $endTime = $startTime.AddMinutes($Minutes)
        $totalSeconds = (New-TimeSpan -Start $startTime -End $endTime).TotalSeconds
        $totalSecondsChild = Get-Random -Minimum 4 -Maximum 30
        $startTimeChild = $startTime
        $endTimeChild = $startTimeChild.AddSeconds($totalSecondsChild)
        $loadingMessage = $loading[(Get-Random -Minimum 0 -Maximum ($loading.Length - 1))]
        #used when progress style is random
        $progcolors = "black", "darkgreen", "magenta", "blue", "darkgray"
    } #begin
    Process {
        #this does not work in VS Code
        if ($host.name -match 'Visual Studio Code') {
            Write-Warning "This command will not work in VS Code."
            #bail out
            Return
        }
        Do {   
            $now = Get-Date
            $secondsElapsed = (New-TimeSpan -Start $startTime -End $now).TotalSeconds
            $secondsRemaining = $totalSeconds - $secondsElapsed
            $percentDone = ($secondsElapsed / $totalSeconds) * 100
            Write-Progress -id 0 -Activity $Title -Status $Message -PercentComplete $percentDone -SecondsRemaining $secondsRemaining
            $secondsElapsedChild = (New-TimeSpan -Start $startTimeChild -End $now).TotalSeconds
            $secondsRemainingChild = $totalSecondsChild - $secondsElapsedChild
            $percentDoneChild = ($secondsElapsedChild / $totalSecondsChild) * 100
            if ($percentDoneChild -le 100) {
                Write-Progress -id 1 -ParentId 0 -Activity $loadingMessage -PercentComplete $percentDoneChild -SecondsRemaining $secondsRemainingChild
            }
            if ($percentDoneChild -ge 100 -and $percentDone -le 98) {
                if ($PSBoundParameters.ContainsKey('ProgressStyle') -AND $PSBoundParameters.Item('ProgressStyle') -eq 'random') {
                    $host.PrivateData.progressBackgroundColor = ($progcolors | Get-Random)
                }
                $totalSecondsChild = Get-Random -Minimum 4 -Maximum 30
                $startTimeChild = $now
                $endTimeChild = $startTimeChild.AddSeconds($totalSecondsChild)
                if ($endTimeChild -gt $endTime) {
                    $endTimeChild = $endTime
                }
                $loadingMessage = $loading[(Get-Random -Minimum 0 -Maximum ($loading.Length - 1))]
            }
            Start-Sleep 0.2
        } Until ($now -ge $endTime)
    } #progress
    End {
        if ($saved) {
            #restore value if it has been changed
            $host.PrivateData.ProgressBackgroundColor = $saved
        }
    } #end
} #end function

<#
*******************************************************************************************************

Function recordCommandLineOptions

.DESCRIPTION

This function writes out all the specified and default paramters for the script.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function recordCommandLineOptions
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes all command line options to the log. ....' -toscreen
	}
	Process 
	{
		Try 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message "DL To Convert" -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $dlToConvert -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message "Ignore Invalid DL Member" -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $ignoreInvalidDLMember -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message "Ignore Invalid Managed By" -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $ignoreInvalidManagedByMember -toscreen
			if ( $groupTypeOverride.Length -eq 0 )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message "Group Type Override NULL" -toscreen
			}
			else 
			{
				Write-LogInfo -LogPath $script:sLogFile -Message "Group Type Override" -toscreen
				Write-LogInfo -LogPath $script:sLogFile -Message $groupTypeOverride -toscreen
			}
			Write-LogInfo -LogPath $script:sLogFile -Message "Covert To Contact" -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $convertToContact -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message "Retain On Premises Settings" -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $retainOnPremisesSettings -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message "Retain O365 Cloud Settings" -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $retainO365CloudOnlySettings -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message "Require Interactive Credentials" -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $requireInteractiveCredentials -toscreen
			
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises credentials were successfully obtained.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises credential were not successfully obtained." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
		}
	}
}

<#
*******************************************************************************************************

Function replicateDomainControllersInbound

.DESCRIPTION

This function replicates domain controllers in the active directory inbound.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function replicateDomainControllersInbound

{
	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function replicateDomainControllersInbound...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Replicates the specified domain controller inbound...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $script:adDomainController -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
	}
	Process 
	{
		Try 
		{
			invoke-command -scriptBlock { repadmin /syncall /A } -Session $script:onPremisesADDomainControllerPowerShellSession
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function replicateDomainControllersInbound...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Successfully replicated the domain controller.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			$error.clear()
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function replicateDomainControllersInbound...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The domain controller could not be replicated - this does not cause the script to abend..." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
	}
}

<#
*******************************************************************************************************

Function setADServerSettings

.DESCRIPTION

This function sets the ad server settings to view entire forest.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function setADServerSettings

{
	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function setADServerSettings...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function sets the ad sever settings to view entire forest....' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
	}
	Process 
	{
		Try 
		{
			set-adServerSettings -viewEntireForest:$TRUE
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function setADServerSettings...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'The ad forest settings were set to view entire forest..' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			$error.clear()
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function setADServerSettings...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The AD forest settings were not set to view entire forest...." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function replicateDomainControllersOutbound

.DESCRIPTION

This function replicates domain controllers in the active directory outbound

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function replicateDomainControllersOutbound

{
	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function replicateDomainControllersOutbound...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Replicates the specified domain controller outbound...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $script:adDomainController -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
	}
	Process 
	{
		Try 
		{
			invoke-command -scriptBlock { repadmin /syncall /APe } -Session $script:onPremisesADDomainControllerPowerShellSession
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function replicateDomainControllersOutbound...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Successfully replicated the domain controller outbound.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			$error.clear()
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function replicateDomainControllersOutbound...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The domain controller could not be replicated - this does not cause the script to abend..." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
	}
}

<#
*******************************************************************************************************

Function createOnPremisesADDomainControllerPowershellSession

.DESCRIPTION

This function creates the on premises AD Domain Controller powershell session.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>
Function createOnPremisesADDomainControllerPowershellSession
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function createOnPremisesADDomainControllerPowershellSession....' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the powershell session to the specified domain controller....' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onPremisesADDomainControllerPowerShellSession = New-PSSession -ComputerName $script:adDomainController -Credential $script:onPremisesCredential -Verbose -Name "ADDomainController"
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function createOnPremisesADDomainControllerPowershellSession....' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to the AD Domain Controller was created successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function createOnPremisesADDomainControllerPowershellSession....' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to AD Domain Controller could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function cleanupSessions

.DESCRIPTION

Removes all powershell sessions created by the script.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function cleanupSessions
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function cleans up all powershell sessions....' -toscreen
	}
	Process 
	{
		Try 
		{
			Get-PSSession | Remove-PSSession
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'All powershell sessions have been cleaned up successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "The powershell sessions could not be cleared - manual removal before restarting required" -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function removeOffice365PowerShellSession

.DESCRIPTION

Removes only the powershell session associated with Office 365.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function removeOffice365PowerShellSession
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function removes the Office 365 powershell sessions....' -toscreen
	}
	Process 
	{
		Try 
		{
			remove-pssession $script:office365PowerShellSession
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'All powershell sessions have been cleaned up successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "The powershell sessions could not be cleared - manual removal before restarting required" -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function removeOnPremisesPowershellSession

.DESCRIPTION

Removes only the powershell session associated with On-Premises.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function removeOnPremisesPowershellSession
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function removes the On-Premsies powershell sessions....' -toscreen
	}
	Process 
	{
		Try 
		{
			remove-pssession $script:onPremisesPowerShellSession
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'All powershell sessions have been cleaned up successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "The powershell sessions could not be cleared - manual removal before restarting required" -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function refreshOffice365PowerShellSession

.DESCRIPTION

Removes, creates, and imports the Office 365 powershell session.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function refreshOffice365PowerShellSession
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function resets the Office 365 powershell sessions....' -toscreen
	}
	Process 
	{
        removeOffice365PowerShellSession  #Removes the Office 365 powershell session.
        createOffice365PowershellSession  #Creates the Office 365 powershell session.
        importOffice365PowershellSession  #Imports the Office 365 powershell session.
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'All Office 365 powershell sessions have been refreshed.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "All Office 365 powershell sessions have not been refreshed." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function refreshOffice365PowerShellSession

.DESCRIPTION

Removes, creates, and imports the Office 365 powershell session.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function refreshOnPremisesPowershellSession
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function resets the Office 365 powershell sessions....' -toscreen
	}
	Process 
	{
        removeOnPremisesPowershellSession  #Removes the Office 365 powershell session.
        createOnPremisesPowershellSession  #Creates the Office 365 powershell session.
        importOnPremisesPowershellSession  #Imports the Office 365 powershell session.
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'All Office 365 powershell sessions have been refreshed.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "All Office 365 powershell sessions have not been refreshed." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}


<#
*******************************************************************************************************

Function establishOnPremisesCredentials

.DESCRIPTION

This function imports the on-premises credentials XML file.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function establishOnPremisesCredentials
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function gathers the on premises secured credentials ....' -toscreen
	}
	Process 
	{
		Try 
		{
			#If the user requires interactive credentials - skip XML and gather credentials.
			
			If ( $requireInteractiveCredentials -eq $FALSE )
			{
				$script:onPremisesCredential = Import-Clixml -Path $script:onPremisesCredentialFile #create the credential variable for local Exchange.
			}
			Else 
			{
				$script:onPremisesCredential = Get-Credential -Message "Please input on premises domain admin and exchange org admin credentials:"
			}
            
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises credentials were successfully obtained.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises credential were not successfully obtained." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function establishOffice365Credentials

.DESCRIPTION

This function imports the on-premises credentials XML file.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function establishOffice365Credentials
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function gathers the Office 365 secured credentials....' -toscreen
	}
	Process 
	{
		Try 
		{
			If ( $requireInteractiveCredentials -eq $FALSE )
			{
				$script:office365Credential = Import-Clixml -Path $script:office365CredentialFile #create the credential variable for Office 365.
			}
			Else
			{
				$script:office365Credential = Get-Credential -Message "Please input global administrator credentials for Office 365:"
			}
            
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The Office 365 credentials were successfully obtained.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The Office 365 credentials were not successfully obtained." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function validateInteractiveCredentials

.DESCRIPTION

This function imports the on-premises credentials XML file.

.PARAMETER credentialToTest

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function validateInteractiveCredentials
{
	Param ($credentialToTest,$passwordType)

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function validates if interactive credentials supplied contain user name and password....' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Testing password type....' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $passwordType -toscreen
		
		$functionCredentialPasswordLength = $credentialToTest.password.Length
		$functionCredentialUserNameLength = $credentialToTest.username.Length
		$functionErrorString = $NULL
	}
	Process 
	{
		Try 
		{
			If ( $functionCredentialUserNameLength -eq 0 )
			{
				$functionErrorString = $passwordType + " Username is blank - cannot proceed"
				Write-Error $functionErrorString -ErrorAction SilentlyContinue
			}
			elseif ( $functionCredentialPasswordLength -eq 0 )
			{
				$functionErrorString = $passwordType + " Password is blank - cannot proceed"
				Write-Error $functionErrorString -ErrorAction SilentlyContinue
			}
			else 
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Username and password combination are not NULL....' -toscreen
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The interactive credentials were successfully validated.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The Office 365 credentials were not successfully obtained." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function createOnPremisesPowershellSession

.DESCRIPTION

This function creates the on premises powershell session to Exchange.
Recommendation to utilize a server Exchange 2016 or newer for avaialbility of all commands.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createOnPremisesPowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the powershell session to on premises Exchange....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onPremisesPowerShellSession = New-PSSession -ConfigurationName $script:onPremisesPowershellConfiguration -ConnectionUri $script:onPremisesPowerShell -Authentication $script:onPremisesPowershellAuthentication -Credential $script:onPremisesCredential -AllowRedirection -Name "ExchangeOnPremises"
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to on premises Exchange was created successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to on premises Exchange could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function createOffice365PowershellSession

.DESCRIPTION

This function creates the office 365 powershell session to Exchange Online.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createOffice365PowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the powershell session to Office 365....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:office365PowerShellSession = New-PSSession -ConfigurationName $script:office365PowershellConfiguration -ConnectionUri $script:office365PowerShell -Authentication $script:office365PowershellAuthentication -Credential $script:office365Credential -AllowRedirection -Name "Office365"
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to Office 365 was created successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to Office 365 could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function createOnPremisesAADConnectPowershellSession

.DESCRIPTION

This function creates the on premises AAD Connect powershell session.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createOnPremisesAADConnectPowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the powershell session to AAD Connect....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onPremisesAADConnectPowerShellSession = New-PSSession -ComputerName $script:onPremisesAADConnectServer -Credential $script:onPremisesCredential -Name "ADConnect"
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to AAD Connect was created successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to AAD Connect could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function importOnPremisesPowershellSession

.DESCRIPTION

This function imports the powershell session to on premises Exchange.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function importOnPremisesPowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function imports the powershell session to on premises Exchange....' -toscreen
	}
	Process 
	{
		Try 
		{
            Import-PSSession $script:onPremisesPowerShellSession
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
            setADServerSettings
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to on premises Exchange was imported successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to on premises Exchange could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function importOffice365PowershellSession

.DESCRIPTION

This function imports the powershell session to Office 365.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function importOffice365PowershellSession
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function imports the powershell session to Office 365....' -toscreen
	}
	Process 
	{
		Try 
		{
            Import-PSSession $script:office365PowerShellSession -Prefix o365
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The powershell session to Office 365 was imported successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The powershell session to Office 365 could not be established - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function collectOnPremsiesDLConfiguration

.DESCRIPTION

This function collects the configuration of the on premises distribution list.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectOnPremsiesDLConfiguration
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collects the on premises distribution list configuration....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onpremisesdlConfiguration = Get-DistributionGroup -identity $dlToConvert -domaincontroller $script:adDomainController
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was collected successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function collectNewOffice365DLInformation

.DESCRIPTION

This function collects the configuration of the Office 365 distribution list.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectNewOffice365DLInformation
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collects the new office 365 distribution list configuration....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:newOffice365DLConfiguration = get-o365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was collected successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function collectNewOffice365DLMemberInformation

.DESCRIPTION

This function collects the configuration of the Office 365 distribution group after creation.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectNewOffice365DLMemberInformation
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collects the new office 365 distribution list member configuration....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:newOffice365DLConfigurationMembership=get-o365DistributionGroupMember -identity $script:onpremisesdlConfiguration.primarySMTPAddress -resultsize unlimited
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was collected successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function collectOffice365DLConfiguation

.DESCRIPTION

This function collects the configuration of the Office 365 distribution list.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectOffice365DLConfiguation
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collects the Office 365 distribution list configuration....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:office365DLConfiguration = Get-o365DistributionGroup -identity $dlToConvert
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The Office 365 distribution list information was collected successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The Office 365 distribution list information could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function performOffice365SafetyCheck

.DESCRIPTION

This funciton reviews the settings of the cloud DL for dir sync status.
If dir sync is not true - assume DL specified exsits on premsies and in the cloud with same address.
This would indicated either direct cloud DL creation <or> migrated previously.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function performOffice365SafetyCheck
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function validates a cloud DLs saftey to migrate....' -toscreen
	}
	Process 
	{
			if ($script:office365DLConfiguration.IsDirSynced -eq $FALSE)
			{
				Write-LogError -LogPath $script:sLogFile -Message 'The DL requested for conversion was found in Office 365 and is not directory synced.  Cannot proceed.'
				Write-Error -Message "ERROR"
			}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The DL is safe to proeced for conversion - source of authority is on-premises.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The DL requested for conversion was found in Office 365 and is not directory synced.  Cannot proceed." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function backupOnPremisesDLConfiguration

.DESCRIPTION

This function writes the configuration of the on premies distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupOnPremisesDLConfiguration
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the on prmeises distribution list configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onpremisesdlConfiguration | Export-CLIXML -Path $script:onPremisesXML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function backupNewOffice365DLConfiguration

.DESCRIPTION

This function writes the configuration of the new office 365 distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupNewOffice365DLConfiguration
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the new Office 365 distribution list configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:newOffice365DLConfiguration | Export-CLIXML -Path $script:newOffice365XML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function backupNewOffice365DLConfigurationMembership

.DESCRIPTION

This function writes the configuration of the new Office 365 distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupNewOffice365DLConfigurationMembership
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the new Office 365 distribution list membership configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:newOffice365DLConfigurationMembership | Export-CLIXML -Path $script:newOffice365MembershipXML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function backupOffice365DLConfiguration

.DESCRIPTION

This function writes the configuration of the Office 365 distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupOffice365DLConfiguration
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the Office 365 distribution list configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:office365DLConfiguration | Export-CLIXML -Path $script:office365XML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The Office 365 distribution list information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The Office 365 distribution list information could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function backupOnPremisesMemberOf

.DESCRIPTION

This function records the groups that the migrated group is a member of.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupOnPremisesMemberOf
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function records the groups that the migrated group is a member of to XML...' -toscreen
	}
	Process 
	{
		Try 
		{
			if ( $script:onPremisesDLMemberOf -ne $NULL )
			{
				$script:onPremisesDLMemberOf | Export-CLIXML -Path $script:onPremisesMemberOfXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises member of for the migrated group has been recorded to XML.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises member of for the migrated group could not be recorded to XML." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function backupOnPremisesMultiValuedAttributes

.DESCRIPTION

This function records the multi-valued attributes of the DLs where permissions are set on premises.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupOnPremisesMultiValuedAttributes
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function records multivalued attributes that the migrated group is a member of to XML...' -toscreen
	}
	Process 
	{
		Try 
		{
			if ( $script:originalGrantSendOnBehalfTo -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing grant send on behalf to to XML...' -toscreen
				$script:originalGrantSendOnBehalfTo | Export-CLIXML -Path $script:originalGrantSendOnBehalfToXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ($script:originalAcceptMessagesFrom -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing accept messages from  to to XML...' -toscreen
				$script:originalAcceptMessagesFrom | Export-CLIXML -Path $script:originalAcceptMessagesFromXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalManagedBy -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing managed by to to XML...' -toscreen
				$script:originalManagedBy | Export-CLIXML -Path $script:originalManagedByXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalRejectMessagesFrom -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing reject messages from to XML...' -toscreen
				$script:originalRejectMessagesFrom | Export-CLIXML -Path $script:originalRejectMessagesFromXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalBypassModerationFromSendersOrMembers -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing bypass mdoeration to XML...' -toscreen
				$script:originalBypassModerationFromSendersOrMembers | Export-CLIXML -Path $script:originalBypassModerationFromSendersOrMembersXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalForwardingAddress -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing forwarding address to XML...' -toscreen
				$script:originalForwardingAddress | Export-CLIXML -Path $script:originalForwardingAddressXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises multivalued attributes for the migrated group has been recorded to XML.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises multivalued attributes for the migrated group could not be recorded to XML." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function backupO365RetainedSettings

.DESCRIPTION

This function records all the values of the cloud only distribution groups that have dependencies on the group migrating.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupO365RetainedSettings
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function exports all Office 365 non-dir sync objects that have dependencies on the migrated group...' -toscreen
	}
	Process 
	{
		Try 
		{
			if ( $script:originalO365GrantSendOnBehalfTo -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing Office 365 grant send on behalf to to XML...' -toscreen
				$script:originalO365GrantSendOnBehalfTo| Export-CLIXML -Path $script:originalO365GrantSendOnBehalfToXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ($script:originalO365AcceptMessagesOnlyFromDLMembers -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing Office 365 accept messages from  to to XML...' -toscreen
				$script:originalO365AcceptMessagesOnlyFromDLMembers | Export-CLIXML -Path $script:originalO365AcceptMessagesFromXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalO365ManagedBy -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing Office 365 managed by to to XML...' -toscreen
				$script:originalO365ManagedBy | Export-CLIXML -Path $script:originalO365ManagedByXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalO365RejectMessagesFromDLMembers -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing Office 365 reject messages from to XML...' -toscreen
				$script:originalO365RejectMessagesFromDLMembers | Export-CLIXML -Path $script:originalO365RejectMessagesFromXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalO365BypassModerationFromSendersOrMembers -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing Office 365 bypass moderation to XML...' -toscreen
				$script:originalO365BypassModerationFromSendersOrMembers | Export-CLIXML -Path $script:originalO365BypassModerationFromSendersOrMembersXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalO365ForwardingAddress -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing Office 365 forwarding address to XML...' -toscreen
				$script:originalO365ForwardingAddress | Export-CLIXML -Path $script:originalO365ForwardingAddressXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalO365MemberOf -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing o365 Member of to a xml file...' -toscreen
				$script:originalO365MemberOf | Export-CLIXML -Path $script:originalO365MemberOfXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalO365GroupGrantSendOnBehalfTo -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing Office 365 Universal Groups Grant Send On Behalf To to XML...' -toscreen
				$script:originalO365GroupGrantSendOnBehalfTo | Export-CLIXML -Path $script:originalO365GroupGrantSendOnBehalfToXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalO365GroupAcceptMessagesOnlyFromDLMembers -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing Office 365 Universal Groups Grant Accept Messages From To to XML...' -toscreen
				$script:originalO365GroupAcceptMessagesOnlyFromDLMembers | Export-CLIXML -Path $script:originalO365GroupAcceptMessagesFromXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:originalO365GroupRejectMessagesFromDLMembers -ne $NULL )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Writing Office 365 Universal Groups Reject Messages From To to XML...' -toscreen
				$script:originalO365GroupRejectMessagesFromDLMembers | Export-CLIXML -Path $script:originalO365GroupRejectMessagesFromXML
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The O365 multivalued attributes for the migrated group has been recorded to XML.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The O365 multivalued attributes for the migrated group could not be recorded to XML." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function backupOnPremisesDLArrays

.DESCRIPTION

This function backs up all the calculated array objects for the DL.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupOnPremisesDLArrays
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function backs up all the calculated array objects for the DL....' -toscreen
	}
	Process 
	{
		Try 
		{
			if ( $script:onpremisesdlconfigurationMembershipArray -ne $NULL )
			{
				$script:onpremisesdlconfigurationMembershipArray | Export-CLIXML -Path $script:onpremisesdlconfigurationMembershipArrayXMLPath
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ($script:onpremisesdlconfigurationManagedByArray -ne $NULL )
			{
				$script:onpremisesdlconfigurationManagedByArray | Export-CLIXML -Path $script:onpremisesdlconfigurationManagedByArrayXMLPath
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:onpremisesdlconfigurationModeratedByArray -ne $NULL )
			{
				$script:onpremisesdlconfigurationModeratedByArray | Export-CLIXML -Path $script:onpremisesdlconfigurationModeratedByArrayXMLPath
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:onpremisesdlconfigurationGrantSendOnBehalfTOArray -ne $NULL )
			{
				$script:onpremisesdlconfigurationGrantSendOnBehalfTOArray | Export-CLIXML -Path $script:onpremisesdlconfigurationGrantSendOnBehalfTOArrayXMLPath
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers -ne $NULL )
			{
				$script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers | Export-CLIXML -Path $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembersXMLPath
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers -ne $NULL )
			{
				$script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers | Export-CLIXML -Path $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembersXMLPath 
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			if ( $script:onPremsiesDLBypassModerationFromSendersOrMembers -ne $NULL )
			{
				$script:onPremsiesDLBypassModerationFromSendersOrMembers | Export-CLIXML -Path $script:onPremsiesDLBypassModerationFromSendersOrMembersXMLPath
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises multivalued attributes for the migrated group has been recorded to XML.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises multivalued attributes for the migrated group could not be recorded to XML." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function archiveFiles

.DESCRIPTION

This function archives the migrated DL files.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function archiveFiles
{
	Param ()

	Begin 
	{
		$functionDate = Get-Date -Format FileDateTime
		$script:archiveXMLPath = $script:onpremisesdlConfiguration.alias + $functionDate
	}
	Process 
	{
		Try 
		{
			rename-item -path $script:sLogPath -newname $script:archiveXMLPath
		}
		Catch 
		{
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-Host "No Error Archive"
		}
		else
		{
			Write-Host "Error"
		}
	}
}

<#
*******************************************************************************************************

Function backupOnPremisesdlMembership

.DESCRIPTION

This function writes the configuration of the on premies distribution list to XML as a backup.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function backupOnPremisesdlMembership
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function writes the on prmeises distribution list membership configuration to XML....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onpremisesdlconfigurationMembership | Export-CLIXML -Path $script:onPremsiesMembershipXML
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The on premises distribution list membership information was written to XML successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The on premises distribution list information membership could not be written to XML - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function collectOnPremisesDLConfigurationMembership

.DESCRIPTION

This function collects the on-premises DL membership.

.PARAMETER <Parameter_Name>

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function collectOnPremisesDLConfigurationMembership
{
	Param ()

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function collections the on premises DL membership....' -toscreen
	}
	Process 
	{
		Try 
		{
            $script:onpremisesdlconfigurationMembership = get-distributionGroupMember -identity $dlToConvert -resultsize unlimited -ignoreDefaultScope
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
            Write-LogInfo -LogPath $script:sLogFile -Message 'The DL membership was collected successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The DL membership could not be collected - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function buildMembershipArray

.DESCRIPTION

This function builds an array of PSObjects representing DL memmbers or multi-valued attributes.

.PARAMETER OperationType

String $operationType - specifies the multi valued attribute we are working with.
String $arrayName - specifies the script variable array name to work with.
String $ignoreVariable - specifies if we should ignore errors found where multi0valued objects have users that are not represented in Office 365

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>
Function buildMembershipArray
{
    Param ([string]$operationType,[string]$arrayName,[boolean]$ignoreVariable=$FALSE)

	Begin 
	{
        [array]$functionArray = @()
		[array]$functionOutput = @()
		$functionRecipient = $NULL
		$functionObject= $NULL
		$recipientObject = $NULL
		$userObject = $NULL

		Write-LogInfo -LogPath $script:sLogFile -Message 'This function builds an array of DL members or multivalued attributes....' -toscreen
	}
	Process 
	{
		#Based on the operation performed move a copy of the script array into a function array.
		#Function array allows us to reuse some code below.

        if ($arrayName -eq "onpremisesdlconfigurationMembershipArray")
        {
            $functionArray = $script:onpremisesdlconfigurationMembership
        }
        elseif ($arrayName -eq "onpremisesdlconfigurationManagedByArray")
        {
			$functionArray = $script:onpremisesdlConfiguration.ManagedBy
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationModeratedByArray")
		{
			$functionArray = $script:onpremisesdlConfiguration.ModeratedBy
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationGrantSendOnBehalfTOArray")
		{
			$functionArray = $script:onpremisesdlConfiguration.GrantSendOnBehalfTo
		}
		elseif ($arrayName -eq  "onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers")
		{
			$functionArray = $script:onpremisesdlConfiguration.AcceptMessagesOnlyFromSendersOrMembers
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationRejectMessagesFromSendersOrMembers")
		{
			$functionArray = $script:onpremisesdlConfiguration.RejectMessagesFromSendersOrMembers
		}
		elseif ($arrayName -eq "onPremsiesDLBypassModerationFromSendersOrMembers")
		{
			$functionArray = $script:onpremisesdlConfiguration.BypassModerationFromSendersOrMembers
		}

		#Based on the operation type passed act on the array of members.
		
        if ( $operationType -eq 'DLMembership' ) 
        {
			#Operation type is distribution list membership.

            foreach ( $member in $functionArray )
            {
				#If the recipient returned is not a user, contact, or groups it is mail recipient.
				#Gather information and place into PSObject.
				#Commit PSobject to output array.

                if ( ($member.recipientType.tostring() -ne "USER") -and ($member.recipientType.tostring() -ne "CONTACT") -and ($member.recipientType.tostring() -ne "GROUP") )
                {
					Write-LogInfo -LogPath $script:sLogFile -Message "Processing mail enabled DL member:" -ToScreen
					Write-LogInfo -LogPath $script:sLogFile -Message $member.name -ToScreen

                    $functionRecipient = get-recipient -identity $member.PrimarySMTPAddress

					if ( $functionRecipient.CustomAttribute1 -eq "MigratedByScript")
					{
						Write-LogInfo -LogPath $script:sLogFile -Message "This member is a migrated DL converted to mail contact." -ToScreen

						$recipientObject = New-Object PSObject -Property @{
							Alias = $functionRecipient.Alias
							Name = $functionRecipient.Name
							PrimarySMTPAddressOrUPN = $functionRecipient.CustomAttribute2
							GUID = $NULL
							RecipientType = "MailUniversalDistributionGroup"
							RecipientOrUser = "Recipient"
						}
					}
					else 
					{
						Write-LogInfo -LogPath $script:sLogFile -Message "This member is a mail enabled user." -ToScreen

						$recipientObject = New-Object PSObject -Property @{
							Alias = $functionRecipient.Alias
							Name = $functionRecipient.Name
							PrimarySMTPAddressOrUPN = $functionRecipient.PrimarySMTPAddress
							GUID = $NULL
							RecipientType = $functionRecipient.RecipientType
							RecipientOrUser = "Recipient"
						}
					}

                    $functionOutput += $recipientObject
				}
                elseif ( $member.recipientType.toString() -eq "USER" )
                {
					Write-LogInfo -LogPath $script:sLogFile -Message "Processing non-mailenabled DL member:" -ToScreen
					Write-LogInfo -LogPath $script:sLogFile -Message $member.name -ToScreen
                    $functionUser = get-user -identity $member.name

					$userObject = New-Object PSObject -Property @{
						Alias = $NULL
						Name = $functionRecipient.Name
						PrimarySMTPAddressOrUPN = $functionUser.UserprincipalName
						GUID = $NULL
						RecipientType = "User"
						RecipientOrUser = "User"
					}
					
                    $functionOutput += $userObject
                }
                elseif ($ignoreVariable -eq $FALSE )
                {
                    Write-LogError -LogPath $script:sLogFile -Message $member.name -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "A non-mail enabled or Office 365 object was found in the group." -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "Script invoked without skipping invalid DL Member." -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "The object must be removed or mail enabled." -ToScreen   
                    Write-LogError -LogPath $script:sLogFile -Message "EXITING." -ToScreen 
                    cleanupSessions
					Stop-Log -LogPath $script:sLogFile -ToScreen
					archiveFiles  
                }
                else 
                {
                    #DO nothing - the user indicated that invalid recipients that cannot be migrated can be skipped.

                    Write-LogInfo -LogPath $script:sLogFile -Message "The following object was intentionally skipped - object type not replicated to Exchange Online" -ToScreen
                    Write-LogInfo -LogPath $script:sLogFile -Message $member.name -ToScreen
                }
            }
		}
		
		#Operation type is managed by array.

        elseif ( $operationType -eq "ManagedBy"  )
        {
            foreach ( $member in $functionArray )
            {
				#Attempt to get a recipient.  If a recipient is returned the object is mail enabled.
				#This approach was taken becuase the attributes are stored in DN format on premises - and may not reflect a mail enabled object.
				#Managed by can be edited directly through AD and may contain non-mail enabled objects.

                if ( Get-Recipient -identity $member -errorAction SilentlyContinue )
                {
					Write-LogInfo -LogPath $script:sLogFile -Message "Processing Managed By member:" -ToScreen
					Write-LogInfo -LogPath $script:sLogFile -Message $member -ToScreen

					$functionRecipient = get-recipient -identity $member

					if ( $functionRecipient.CustomAttribute1 -eq "MigratedByScript")
					{
						Write-LogInfo -LogPath $script:sLogFile -Message "This member is a migrated DL converted to mail contact." -ToScreen

						$recipientObject = New-Object PSObject -Property @{
							Alias = $functionRecipient.Alias
							Name = $functionRecipient.Name
							PrimarySMTPAddressOrUPN = $functionRecipient.CustomAttribute2
							GUID = $NULL
							RecipientType = "MailUniversalDistributionGroup"
							RecipientOrUser = "Recipient"
						}
					}
					else 
					{
						Write-LogInfo -LogPath $script:sLogFile -Message "This member is a mail enabled user." -ToScreen

						$recipientObject = New-Object PSObject -Property @{
							Alias = $functionRecipient.Alias
							Name = $functionRecipient.Name
							PrimarySMTPAddressOrUPN = $functionRecipient.PrimarySMTPAddress
							GUID = $NULL
							RecipientType = $functionRecipient.RecipientType
							RecipientOrUser = "Recipient"
						}
					}

					$functionOutput += $recipientObject
                }
                elseif ($ignoreVariable -eq $FALSE )
                {
                    Write-LogError -LogPath $script:sLogFile -Message $member -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "A non-mail enabled or Office 365 object was found in ManagedBy." -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "Script invoked without skipping invalid DL Member." -ToScreen
                    Write-LogError -LogPath $script:sLogFile -Message "The object must be removed or mail enabled." -ToScreen   
                    Write-LogError -LogPath $script:sLogFile -Message "EXITING." -ToScreen 
                    cleanupSessions
					Stop-Log -LogPath $script:sLogFile -ToScreen  
					archiveFiles
                }
                else 
                {
                    #DO nothing - the user indicated that invalid recipients that cannot be migrated can be skipped.

                    Write-LogInfo -LogPath $script:sLogFile -Message "The following object was intentionally skipped - object type not replicated to Exchange Online" -ToScreen
                    Write-LogInfo -LogPath $script:sLogFile -Message $member -ToScreen
                }
            }
		}

		#Operation is a remaining multivalued attribute.

		elseif ( ( $operationType -eq "ModeratedBy" ) -or ( $operationType  -eq "GrantSendOnBehalfTo" ) -or ( $operationType -eq "AcceptMessagesOnlyFromSendersOrMembers") -or ($operationType -eq "RejectMessagesFromSendersOrMembers" ) -or ($operationType -eq "BypassModerationFromSendersOrMembers") )
		{
			foreach ( $member in $functionArray )
            {
				#Test to ensure that the object is a recipient.
				#In theory this is not required since Exchange commandlets will not let you set a non-mail enabled object on these properties...but...
				#For consistency sake we're doing the same thing...

                if ( Get-Recipient -identity $member -errorAction SilentlyContinue )
                {
					Write-LogInfo -LogPath $script:sLogFile -Message "Processing ModeratedBy, GrantSendOnBehalfTo, AcceptMessagesOnlyFromSendersorMembers, RejectMessagesFromSendersOrMembers, or BypassModerationFromSendersOrMembers member:" -ToScreen
					Write-LogInfo -LogPath $script:sLogFile -Message $member -ToScreen

					$functionRecipient = get-recipient -identity $member

					if ( $functionRecipient.CustomAttribute1 -eq "MigratedByScript")
					{
						Write-LogInfo -LogPath $script:sLogFile -Message "This member is a migrated DL converted to mail contact." -ToScreen

						$recipientObject = New-Object PSObject -Property @{
							Alias = $functionRecipient.Alias
							Name = $functionRecipient.Name
							PrimarySMTPAddressOrUPN = $functionRecipient.CustomAttribute2
							GUID = $NULL
							RecipientType = "MailUniversalDistributionGroup"
							RecipientOrUser = "Recipient"
						}
					}
					else 
					{
						Write-LogInfo -LogPath $script:sLogFile -Message "This member is a mail enabled user." -ToScreen

						$recipientObject = New-Object PSObject -Property @{
							Alias = $functionRecipient.Alias
							Name = $functionRecipient.Name
							PrimarySMTPAddressOrUPN = $functionRecipient.PrimarySMTPAddress
							GUID=$NULL
							RecipientType = $functionRecipient.RecipientType
							RecipientOrUser = "Recipient"
						}
					}

					$functionOutput += $recipientObject
                }
			}
		}

		#Based on the array name iterate and log the items found to process.

        if ( $arrayName -eq "onpremisesdlconfigurationMembershipArray" )
        {
			$script:onpremisesdlconfigurationMembershipArray = $functionOutput

			foreach ( $member in $script:onpremisesdlconfigurationMembershipArray )
            {
                Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
        }
        elseif ($arrayName -eq "onpremisesdlconfigurationManagedByArray")
        {
			$script:onpremisesdlconfigurationManagedByArray = $functionOutput
			
			foreach ( $member in $script:onpremisesdlconfigurationManagedByArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationModeratedByArray")
		{
			$script:onpremisesdlconfigurationModeratedByArray = $functionOutput
			
			foreach ( $member in $script:onpremisesdlconfigurationModeratedByArray )
            {
                Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationGrantSendOnBehalfTOArray")
		{
			$script:onpremisesdlconfigurationGrantSendOnBehalfTOArray = $functionOutput

			foreach ( $member in $script:onpremisesdlconfigurationGrantSendOnBehalfTOArray )
            {
                Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers")
		{
			$script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers = $functionOutput

			foreach ( $member in $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers )
            {
                Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
                Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen 
            }
		}
		elseif ($arrayName -eq "onpremisesdlconfigurationRejectMessagesFromSendersOrMembers")
		{
			$script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers = $functionOutput

			foreach ( $member in $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers)
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
				Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen
			}
		}
		elseif ($arrayName -eq "onPremsiesDLBypassModerationFromSendersOrMembers")
		{
			$script:onPremsiesDLBypassModerationFromSendersOrMembers = $functionOutput

			foreach ( $member in $script:onPremsiesDLBypassModerationFromSendersOrMembers)
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'The following SMTP address was added to the array:' -ToScreen
				Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -ToScreen
			}
		}
	}
	End 
	{
		If ($?) 
		{
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message 'The array was built successfully.' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The array could not be built successfully - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function testOffice365Recipient

.DESCRIPTION

This function collects the on-premises DL membership.

.PARAMETER <Parameter_Name>

[string]$primarySMTPAddressOrUPN is the UPN or Primary SMTP address we are testing office 365 for.
[string]$userOrRecipient determines if we test with get-recipient (recipient) or get-user (user)

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function testOffice365Recipient
{
	Param ([string]$primarySMTPAddressOrUPN,[string]$UserorRecipient)

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function validates that all objects in the passed array exist in Office 365....' -toscreen
	}
	Process 
	{
		Try 
		{
			#Test to ensure the mail enabled or user object exists in Office 365.

			Write-LogInfo -LogPath $script:sLogFile -Message "Testing user in Office 365..." -ToScreen
			Write-LogInfo -LogPath $script:sLogFile -Message $primarySMTPAddressOrUPN -ToScreen
			Write-LogInfo -LogPath $script:sLogFile -Message $UserorRecipient -toScreen
			
			if ( $UserorRecipient -eq "Recipient")
			{
				$fixedPrimarySMTPAddressOrUPN = "$($primarySMTPAddressorUPN.replace("'","''"))"
				$functionCommand = "get-o365recipient -filter {primarySMTPAddress -eq '$fixedPrimarySMTPAddressOrUPN'}"

				$functionTest = Invoke-Expression $functionCommand

				if ( !$functionTest )
				{
					throw ("User or recipient not found in office 365 - all recipients and users must be in Office 365 - " + $primarySMTPAddress )
				}
            	
				Write-LogInfo -LogPath $script:sLogFile -Message $functionTest.GUID -toScreen
				$script:arrayGUID = $functionTest.GUID.tostring()
                Write-LogInfo -LogPath $script:sLogFile -Message $script:arrayGUID -toScreen
			}
			elseif ($UserorRecipient -eq "User")
			{
				$fixedPrimarySMTPAddressorUPN = "$($primarySMTPAddressOrUPN.Replace("'","''"))"
				$functionCommand = "get-o365User -filter {userPrincipalName -eq '$fixedPrimarySMTPAddressorUPN'}"

				$functionTest=invoke-expression $functionCommand

				if ( !$functionTest )
				{
					throw ("User or recipient not found in office 365 - all recipients and users must be in Office 365 - " + $primarySMTPAddress )
				}

				Write-LogInfo -LogPath $script:sLogFile -Message $functionTest.GUID -ToScreen
				$script:arrayGUID = $functionTest.GUID.tostring()
                Write-LogInfo -LogPath $script:sLogFile -Message $script:arrayGUID -toScreen
			}
		}
		Catch 
		{
            Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
            cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The recipients were found in Office 365.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			$error.clear()
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The recipients were not found in Office 365 - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function testOffice365GroupMigrated

.DESCRIPTION

This function tests to ensure that a sub group or permissions groups was migrated before the DL that references it.

If we did not do this migration of sub groups woudl be possible - membership and permissions would be lost.

.PARAMETER <Parameter_Name>

[string]$primarySMTPAddressOrUPN - group reference to test for.

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function testOffice365GroupMigrated
{
	Param ([string]$primarySMTPAddress)

	Begin 
	{
	    Write-LogInfo -LogPath $script:sLogFile -Message 'This function tests to see if any sub groups or groups assigned permissions have been migrated....' -toscreen
	}
	Process 
	{
		#Test the dir sync flag to see if TRUE.  If TRUE the list was not migrated - abend.

		$functionTest = Get-o365DistributionGroup -identity $primarySMTPAddress

		Write-LogInfo -LogPath $script:sLogFile -Message 'Now testing group...' -ToScreen
		Write-LogInfo -LogPath $script:sLogFile -Message $functionTest.primarySMTPAddress -ToScreen
		Write-LogInfo -LogPath $script:sLogFile -Message $functionTest.IsDirSynced -ToScreen

		if ( $functionTest.primarySMTPAddress -eq $script:onpremisesdlConfiguration.primarySMTPAddress )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The group has permissions set to itself - this is alloed to proceed...' -ToScreen
		}
		elseif ( $functionTest.IsDirSynced -eq $TRUE)
		{
			Write-LogError -LogPath $script:sLogFile -Message 'A distribution list was found as a sub-member or on a multi-valued attribute.' -ToScreen
			Write-LogError -LogPath $script:sLogFile -Message 'The distribution list has not been migrated to Office 365 (DirSync Flag is TRUE)' -ToScreen
			Write-LogError -LogPath $script:sLogFile -Message 'All sub lists or lists with permissions must be migrated before proceeding.' -ToScreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The recipients were found in Office 365.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			$error.clear()
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The recipients were not found in Office 365 - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function moveGroupToOU

.DESCRIPTION

This function moves the group to an OU that is not synchonized by AAD Connect

This requires prior configuration of this setting in AAD Connect.

This requires the OU exist prior to script execution.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function moveGroupToOU
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function moves group to the non-sync OU' -toscreen
	}
	Process 
	{
		Try 
		{
			Invoke-Command -ScriptBlock { Move-ADObject -identity $args[0] -TargetPath $args[1] } -ArgumentList $script:onpremisesdlConfiguration.distinguishedName,$script:groupOrganizationalUnit -Session $script:onPremisesADDomainControllerPowerShellSession
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The group has been moved successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "Group Could Not Be Moved...exiting" -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function replicateDomainControllers

.DESCRIPTION

This function replicates domain controllers in the active directory.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function replicateDomainControllers

{
	Param ([string]$domainController,[string]$distinguishedName)

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'Replicates the specified domain controller...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $domainControllerName -toscreen
	}
	Process 
	{
		Try 
		{
			replicateDomainControllersInbound
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			replicateDomainControllersOutbound
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Successfully replicated the domain controller.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			$error.clear()
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "The domain controller could not be replicated - this does not cause the script to abend..." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
		}
	}
}


<#
*******************************************************************************************************

Function invokeADConnect

.DESCRIPTION

This function invokes a delta sync through remote powershell.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function invokeADConnect
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function triggers the ad connect process to sync changes...' -toscreen
	}
	Process 
	{
		Try 
		{
			Invoke-Command -Session $script:onPremisesAADConnectPowerShellSession -ScriptBlock {Import-Module -Name 'AdSync'}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			Invoke-Command -Session $script:onPremisesAADConnectPowerShellSession	-ScriptBlock {start-adsyncsynccycle -policyType Delta}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The AD Connect instance has been successfully initiated.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			$error.clear()
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "ADConnect sync could not be triggered...this does not cause the script to abend." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
		}
	}
}

<#
*******************************************************************************************************

Function createOffice365DistributionList

.DESCRIPTION

This function creates the new cloud DL with the minimum attributes. 

Detailed attributes will be handeled later.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createOffice365DistributionList
{
	Param ()
	
	Begin 
	{
		$functionGroupType = $NULL #Utilized to establish group type distribution or group type security.

		Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the cloud DL with the minimum settings...' -toscreen

		if ( $groupTypeOverride -eq "Security" )
		{
			$functionGroupType="Security"
		}
		elseif ( $groupTypeOverride -eq "Distribution" )
		{
			$functionGroupType="Distribution"
		}
		elseif ( $script:onpremisesdlConfiguration.GroupType -eq "Universal, SecurityEnabled" )
		{
			$functionGroupType="Security"
		}
		elseif (  $script:onpremisesdlConfiguration.GroupType -eq "Universal" )
		{
			$functionGroupType="Distribution"
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "This script only supports universal distribution and universal security group conversions." -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
		$script:onpremisesdlConfiguration.GroupType
	}
	Process 
	{
		Try 
		{
			new-o365DistributionGroup -name $script:onpremisesdlConfiguration.Name -alias $script:onpremisesdlConfiguration.Alias -type $functionGroupType
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution list created successfully in Exchange Online / Office 365.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "Distribution list could not be created in Exchange Online / Office 365...exiting" -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function setOffice365DistributionListSettings

.DESCRIPTION

This function sets the single attribute settings on the new Cloud DL based on the previous on-premises DL.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function setOffice365DistributionListSettings
{
	Param ()

	Begin 
	{
		$functionEmailAddresses = $NULL
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function updates the cloud DL settings to match on premise...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This does not update the multivalued attributes...' -ToScreen

		#Build the X500 address for the new email address based off on premises values.

		$script:x500Address = "X500:"+$script:onpremisesdlConfiguration.legacyExchangeDN
		$functionEmailAddresses=$script:onpremisesdlconfiguration.emailAddresses
		$functionEmailAddresses+=$script:x500Address

		#If the override to security group is established - and the type is security - the member join restrictions must be overridden to Closed.

		if ( $groupTypeOverride -eq "Security" )
		{
			$functionMemberDepartRestriction = "Closed"
		}
		else 
		{
			$functionMemberDepartRestriction = $script:onpremisesdlconfiguration.MemberDepartRestriction
		}
	}
	Process 
	{
		Try 
		{
			Set-O365DistributionGroup -Identity $script:onpremisesdlConfiguration.alias -BypassNestedModerationEnabled $script:onpremisesdlconfiguration.BypassNestedModerationEnabled -MemberJoinRestriction $script:onpremisesdlconfiguration.MemberJoinRestriction -MemberDepartRestriction $functionMemberDepartRestriction -ReportToManagerEnabled $script:onpremisesdlconfiguration.ReportToManagerEnabled -ReportToOriginatorEnabled $script:onpremisesdlconfiguration.ReportToOriginatorEnabled -SendOofMessageToOriginatorEnabled $script:onpremisesdlconfiguration.SendOofMessageToOriginatorEnabled -Alias $script:onpremisesdlconfiguration.Alias -CustomAttribute1 $script:onpremisesdlconfiguration.CustomAttribute1 -CustomAttribute10 $script:onpremisesdlconfiguration.CustomAttribute10 -CustomAttribute11 $script:onpremisesdlconfiguration.CustomAttribute11 -CustomAttribute12 $script:onpremisesdlconfiguration.CustomAttribute12 -CustomAttribute13 $script:onpremisesdlconfiguration.CustomAttribute13 -CustomAttribute14 $script:onpremisesdlconfiguration.CustomAttribute14 -CustomAttribute15 $script:onpremisesdlconfiguration.CustomAttribute15 -CustomAttribute2 $script:onpremisesdlconfiguration.CustomAttribute2 -CustomAttribute3 $script:onpremisesdlconfiguration.CustomAttribute3 -CustomAttribute4 $script:onpremisesdlconfiguration.CustomAttribute4 -CustomAttribute5 $script:onpremisesdlconfiguration.CustomAttribute5 -CustomAttribute6 $script:onpremisesdlconfiguration.CustomAttribute6 -CustomAttribute7 $script:onpremisesdlconfiguration.CustomAttribute7 -CustomAttribute8 $script:onpremisesdlconfiguration.CustomAttribute8 -CustomAttribute9 $script:onpremisesdlconfiguration.CustomAttribute9 -ExtensionCustomAttribute1 $script:onpremisesdlconfiguration.ExtensionCustomAttribute1 -ExtensionCustomAttribute2 $script:onpremisesdlconfiguration.ExtensionCustomAttribute2 -ExtensionCustomAttribute3 $script:onpremisesdlconfiguration.ExtensionCustomAttribute3 -ExtensionCustomAttribute4 $script:onpremisesdlconfiguration.ExtensionCustomAttribute4 -ExtensionCustomAttribute5 $script:onpremisesdlconfiguration.ExtensionCustomAttribute5 -DisplayName $script:onpremisesdlconfiguration.DisplayName -HiddenFromAddressListsEnabled $script:onpremisesdlconfiguration.HiddenFromAddressListsEnabled -ModerationEnabled $script:onpremisesdlconfiguration.ModerationEnabled -RequireSenderAuthenticationEnabled $script:onpremisesdlconfiguration.RequireSenderAuthenticationEnabled -SimpleDisplayName $script:onpremisesdlconfiguration.SimpleDisplayName -SendModerationNotifications $script:onpremisesdlconfiguration.SendModerationNotifications -WindowsEmailAddress $script:onpremisesdlconfiguration.WindowsEmailAddress -MailTipTranslations $script:onpremisesdlconfiguration.MailTipTranslations -Name $script:onpremisesdlconfiguration.Name
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		try 
		{
			write-LogInfo -LogPath $script:sLogFile -message "Processing primary proxy address.." -ToScreen
			Write-LogInfo -logPath $script:sLogFile -message $script:onpremisesdlConfiguration.primarySMTPAddress -ToScreen
			set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.alias -primarySMTPAddress $script:onpremisesdlConfiguration.primarySMTPAddress
		}
		catch
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		try
		{
			foreach ( $address in $functionEmailAddresses )
			{
				Write-LogInfo -LogPath $script:sLogFile -message "Processing email address.." -ToScreen
				Write-LogInfo -logPath $script:sLogFile -message $address -ToScreen
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -EmailAddresses @{add=$address}
			}
		}
		catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution group properties updated successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "Cannot update properties of distribution group." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function removeOnPremisesDistributionGroup

.DESCRIPTION

This function removes the on premises distribution group.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function removeOnPremisesDistributionGroup
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function removes the on premises distribution group in preparation for contact conversion' -toscreen
	}
	Process 
	{
		Try 
		{
			remove-DistributionGroup -Identity $script:onpremisesdlConfiguration.primarySMTPAddress -confirm:$FALSE -byPassSecurityGroupManagerCheck:$True -domaincontroller $script:adDomainController
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution group successfully removed.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message "Distribution group could not be successfully removed." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function setOffice365DistributionListSettings

.DESCRIPTION

This function sets the multi valued attribute settings on the new Cloud DL based on the previous on-premises DL.

.PARAMETER 

[string]$operationType is the type of operation - for example moderateydBy or managedBy.
[string]$primarySMTPAddressOrUPN is the primary SMTP address of a recipient or UPN of the user object that we are operating on.

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function setOffice365DistributionlistMultivaluedAttributes
{
	Param ([string]$operationType,[string]$primarySMTPAddressOrUPN)

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function sets the multi-valued attributes' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $operationType -ToScreen
		Write-LogInfo -LogPath $script:sLogFile -Message $primarySMTPAddressOrUPN
	}
	Process 
	{
		#Based on the operation type specified utilize the appropriate array and iterate through each member adding to the attribute in the service.

		if ( $operationType -eq "DLMembership")
		{
			Try
			{
				add-o365DistributionGroupMember -identity $script:onpremisesdlConfiguration.primarySMTPAddress -member $PrimarySMTPAddressOrUPN
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				archiveFiles
				Break
			}
		}
		elseif ( $operationType -eq "ManagedBy")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -managedBy @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				archiveFiles
				Break
			}
		}
		elseif ( $operationType -eq "ModeratedBy")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -ModeratedBy @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				archiveFiles
				Break
			}
		}
		elseif ( $operationType -eq "GrantSendOnBehalfTo")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -GrantSendOnBehalfTo @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				archiveFiles
				Break
			}
		}
		elseif ( $operationType -eq "AcceptMessagesOnlyFromSendersOrMembers")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -AcceptMessagesOnlyFromSendersOrMembers @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				archiveFiles
				Break
			}
		}
		elseif ( $operationType -eq "RejectMessagesFromSendersOrMembers")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -RejectMessagesFromSendersOrMembers @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				archiveFiles
				Break
			}
		}
		elseif ( $operationType -eq "BypassModerationFromSendersOrMembers")
		{
			Try
			{
				set-O365DistributionGroup -identity $script:onpremisesdlConfiguration.primarySMTPAddress -BypassModerationFromSendersOrMembers @{add=$PrimarySMTPAddressOrUPN}
			}
			Catch
			{
				Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
				cleanupSessions
				Stop-Log -LogPath $script:sLogFile -ToScreen
				archiveFiles
				Break
			}
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'The mutilvalued attribute was updated successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $operationType -ToScreen
            Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
		}
		else
		{

			Write-LogError -LogPath $script:sLogFile -Message "The mutilvalued attribute could not be updated successfully - exiting." -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message $operationType -ToScreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function recordDistributionGroupMembership

.DESCRIPTION

Records the memberhsip of the distribution group in the event of mail contact conversion.
This allows us to add the mail enabled contact to the groups on premises.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function recordDistributionGroupMembership
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function recordDistributionGroupMembership...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function records the membership of the on premises distribution group.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
	}
	Process 
	{
		Try 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Invoking AD call to domain controller to pull membership of the group on premises...' -toscreen

			$script:onPremisesDLMemberOf = Invoke-Command -ScriptBlock { get-ADPrincipalGroupMembership -identity $args[0] } -ArgumentList $script:onPremisesMovedDLConfiguration.samAccountName -Session $script:onPremisesADDomainControllerPowerShellSession
		
			foreach ( $member in $script:onPremisesDLMemberOf )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Member Found:' -ToScreen
				Write-LogInfo -LogPath $script:sLogFile -Message $member.distinguishedName -ToScreen
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function recordDistributionGroupMembership...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution group member of successfully capture.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function recordDistributionGroupMembership...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "Distribution group member of could not be captured." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function recordMovedOriginalDistributionGroupProperties

.DESCRIPTION

The original on premises distribution group has been moved.  We need to update the properties so that we can search for it moving forward.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function recordMovedOriginalDistributionGroupProperties
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function recordMovedOriginalDistributionGroupProperties...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function records the properties of the distribution group after the OU has changed.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
	}
	Process 
	{
		Try 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Recording moved DL configuration...' -toscreen

			$script:onPremisesMovedDLConfiguration = get-distributionGroup -identity $dlToConvert -domaincontroller $script:adDomainController
		
			Write-LogInfo -LogPath $script:sLogFile -Message $script:onPremisesMovedDLConfiguration.identity
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function recordMovedOriginalDistributionGroupProperties...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution group successfully capture.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function recordMovedOriginalDistributionGroupProperties...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "Distribution group could not be captured." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function recordOriginalMultivaluedAttributes

.DESCRIPTION

Records the original multi-valued attributes so that we can restamp the contacts back.
The purpose is to have these retained for future DL migrations to the service.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function recordOriginalMultivaluedAttributes
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function recordOriginalMultivaluedAttributes...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Recording the original multi-valued attributes.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen

		#Record the identity of the moved distribution list (updated post move so we have correct identity)
		
		$functionGroupIdentity = $script:onPremisesMovedDLConfiguration.identity.tostring()	#Function variable to hold the identity of the group.
		$fixedFunctionGroupIdentity = "$($functionGroupIdentity.Replace("'","''"))"
		$functionCommand = $NULL	#Holds the expression that we will be executing to determine multi-valued membership.
		[array]$functionGroupArray = @()
		$functionRecipientObject = $NULL
		
		Write-LogInfo -LogPath $script:sLogFile -Message 'The following group identity is the filtered name.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $functionGroupIdentity -toscreen
	}
	Process 
	{
		Try 
		{
			#Using a filter detemrine all groups this group had grant send on behalf to.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all grantSendOnBehalfTo for the identity...' -toscreen

            $functionCommand = "get-distributionGroup -resultsize unlimited -Filter { GrantSendOnBehalfTo -eq '$fixedFunctionGroupIdentity' } -domainController '$script:adDomainController'"
            
            $script:originalGrantSendOnBehalfTo = Invoke-Expression $functionCommand
		
			foreach ( $member in $script:originalGrantSendOnBehalfTo )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message $member.primarySMTPAddress -ToScreen
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Using a filter detemrine all groups this group had accept message from DL members

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all AcceptMessagesOnlyFromDLMembers for the identity...' -toscreen

            $functionCommand = "get-distributionGroup -resultsize unlimited -Filter { AcceptMessagesOnlyFromDLMembers -eq '$fixedFunctionGroupIdentity' } -domainController '$script:adDomainController'"
            
            $script:originalAcceptMessagesFrom = Invoke-Expression $functionCommand
		
			foreach ( $member in $script:originalAcceptMessagesFrom )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message $member.primarySMTPAddress -ToScreen
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Using a filter detemrine all groups this group had managedBy from DL members

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all ManagedBy for the identity...' -toscreen

            $functionCommand = "get-distributionGroup -resultsize unlimited -Filter { ManagedBy -eq '$fixedFunctionGroupIdentity' } -domainController '$script:adDomainController'"
            
            $script:originalManagedBy = Invoke-Expression $functionCommand
		
			foreach ( $member in $script:originalManagedBy )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message $member.primarySMTPAddress -ToScreen
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Using a filter detemrine all groups this group had accept message from DL members

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all RejectMessagesFromDLMembers for the identity...' -toscreen

            $functionCommand = "get-distributionGroup -resultsize unlimited -Filter { RejectMessagesFromDLMembers -eq '$fixedFunctionGroupIdentity' } -domainController '$script:adDomainController'"
            
            $script:originalRejectMessagesFrom = Invoke-Expression $functionCommand
		
			foreach ( $member in $script:originalRejectMessagesFrom )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message $member.primarySMTPAddress -ToScreen
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Perform a server side search of all groups where this group was set to bypass moderation.
			#Note:  There is no filerable attribute for this - this operation can be very expensive, memory intensive, and time consuming.
			#Note:  Administrators may consider commenting out these portions and not attempting to perserve this for other DL migrations.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all bypass moderations for the identity...' -toscreen

			$functionGroupArray = invoke-command -scriptBlock { get-distributiongroup -resultsize unlimited -domainController $script:adDomainController | where { $_.bypassModerationFromSendersOrMembers -eq $fixedFunctionGroupIdentity } }

			foreach ( $loopGroup in $functionGroupArray)
			{
				Write-LogInfo -LogPath $script:sLogFile -Message $loopGroup.primarySMTPAddress -ToScreen

				#Create a custom object of each of the DLs found for later use.

				$functionRecipientObject = New-Object PSObject -Property @{
					DistinguishedName = $loopgroup.distinguishedName
					Alias = $loopGroup.Alias
					Name = $loopGroup.Name
					PrimarySMTPAddressOrUPN = $loopGroup.primarySMTPAddress
				}

				$script:originalBypassModerationFromSendersOrMembers+=$functionRecipientObject		
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Using a filter to determine all mailboxes that have forwardingAddress set to the distribution group.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all forwarding addresses for the identity...' -toscreen

            $functionCommand = "get-mailbox -resultsize unlimited -Filter { ForwardingAddress -eq '$fixedFunctionGroupIdentity' } -domainController '$script:adDomainController'"
            
            $script:originalForwardingAddress = Invoke-Expression $functionCommand
		
			foreach ( $member in $script:originalForwardingAddress )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message $member.primarySMTPAddress -ToScreen
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Perform a server side search of all groups where this group was set to bypass moderation.
			#Note:  There is no filerable attribute for this - this operation can be very expensive, memory intensive, and time consuming.
			#Note:  Administrators may consider commenting out these portions and not attempting to perserve this for other DL migrations.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all bypass moderations for the identity...' -toscreen

			$functionGroupArray = invoke-command -scriptBlock { get-distributiongroup -resultsize unlimited -domainController $script:adDomainController | where { $_.bypassModerationFromSendersOrMembers -eq $fixedFunctionGroupIdentity } }

			foreach ( $loopGroup in $functionGroupArray)
			{
				Write-LogInfo -LogPath $script:sLogFile -Message $loopGroup.primarySMTPAddress -ToScreen

				#Create a custom object of each of the DLs found for later use.

				$functionRecipientObject = New-Object PSObject -Property @{
					DistinguishedName = $loopgroup.distinguishedName
					Alias = $loopGroup.Alias
					Name = $loopGroup.Name
					PrimarySMTPAddressOrUPN = $loopGroup.primarySMTPAddress
				}

				$script:originalBypassModerationFromSendersOrMembers+=$functionRecipientObject		
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
}

<#
*******************************************************************************************************

Function removeOnPremisesDistributionGroup

.DESCRIPTION

This function removes the on premises distribution group.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function removeOnPremisesDistributionGroup
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function removeOnPremisesDistributionGroup...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function removes the on premises distribution group in preparation for contact conversion' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
	}
	Process 
	{
		Try 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Removing on premises distribution group...' -toscreen

			remove-distributionGroup -identity $script:onPremisesMovedDLConfiguration.primarySMTPAddress -domaincontroller $script:adDomainController -confirm:$FALSE -bypassSecurityGroupManagerCheck:$TRUE
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function removeOnPremisesDistributionGroup...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution group successfully removed.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function removeOnPremisesDistributionGroup...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "Distribution group could not be successfully removed." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function createOnPremisesDynamicDistributionGroup

.DESCRIPTION

This function creates a dynamic distribution group that matches the orginal groups information.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createOnPremisesDynamicDistributionGroup
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function createOnPremisesDynamicDistributionGroup...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates the on premsies mail enabled contact to replace the distribution group.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen

		$functionEmailAddressSplit = $script:onpremisesdlConfiguration.primarySMTPAddress.split("@")
		$script:newDynamicDLAddress = $functionEmailAddressSplit[0]+"-Dynamic"+"@"+$functionEmailAddressSplit[1]

		Write-LogInfo -LogPath $script:sLogFile -Message "The identitied dynamic DL email address..." -ToScreen
		Write-LogInfo -LogPath $script:sLogFile -Message $script:newDynamicDLAddress -ToScreen
	}
	Process 
	{
		Try 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Creating dynamic distribution group...' -toscreen

			new-dynamicDistributionGroup -name $script:onpremisesdlConfiguration.name -Alias $script:onpremisesdlconfiguration.Alias -primarySMTPAddress $script:newDynamicDLAddress -organizationalUnit $script:groupOrganizationalUnit -domainController $script:adDomainController -includedRecipients AllRecipients -conditionalCustomAttribute2 $script:onpremisesdlConfiguration.primarySMTPAddress -DisplayName $script:onpremisesdlconfiguration.DisplayName
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function createOnPremisesDynamicDistributionGroup...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'The contact was created successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function createOnPremisesDynamicDistributionGroup...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The contact was not created successfully." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function setOnPremisesDynamicDistributionGroupSettings

.DESCRIPTION

This function mirrors the original distribution list settings on the newly created dynamic distribution group.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function setOnPremisesDynamicDistributionGroupSettings
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function setOnPremisesDynamicDistributionGroupSettings...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function sets the properties of the on-premises dynamic distribution group.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen

		$functionEmailAddresses = $NULL	#Utilized to hold working email addresses in the function.
		$functionRemoteRoutingAddress = $NULL #Utilized to hold the remote routing address found on the original group.

		#Add the original DLs X500 address to the list of proxy addresses.

		$functionEmailAddresses = $script:onpremisesdlConfiguration.emailaddresses
		$functionEmailAddresses+=$script:x500Address

		#Iterate through all proxy addresses to find the remote routing address.
		#This needs to be removed so that it can be stamped on the mail contact matching this group.
		
		foreach ( $emailAddress in $functionEmailAddresses)
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Iterating through proxy addresses to find remote routing address...' -toscreen

			if ( $emailAddress -like "*.mail.onmicrosoft.com" )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'The remote routing address has been found...' -toscreen

				$script:remoteRoutingAddress=$emailAddress
			}
		}
	}
	Process 
	{
		Try 
		{
			#Create the dynamic distribution list where custom attribute 1 equals MigratedByGroup.  Utilize the same information as the original list.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Apply the settings to the dynamic distribution group...' -toscreen

			set-dynamicDistributionGroup -identity $script:newDynamicDLAddress -primarySMTPAddress $script:newDynamicDLAddress -HiddenFromAddressListsEnabled $script:onpremisesdlconfiguration.HiddenFromAddressListsEnabled -SimpleDisplayName $script:onpremisesdlconfiguration.SimpleDisplayName -WindowsEmailAddress $script:newDynamicDLAddress -Name $script:onpremisesdlconfiguration.Name -domaincontroller $script:adDomainController -RequireSenderAuthenticationEnabled $FALSE
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Adding each email address from the original DL to the new dynamic DL.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Adding original proxy addresses to the dynamic DL...' -toscreen

			foreach ( $address in $functionEmailAddresses )
			{
				$address = $address.tolower()
				set-dynamicDistributionGroup -identity $script:newDynamicDLAddress -EmailAddresses @{add=$address} -domaincontroller $script:adDomainController
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Remote the remote routing address from the dynamic distribution lists proxy addresses.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Remvoe the remote routing address from the list of proxy addresses...' -toscreen

			set-dynamicDistributionGroup -identity $script:newDynamicDLAddress -EmailAddresses @{remove=$script:remoteRoutingAddress}  -domaincontroller $script:adDomainController
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Set the OU filter for the group to search for the mail contact in the same OU as where the origianl group resided.
			#This is required becuase when creating a dynamicDL it sets the container filter to match OU - which will not work in this case.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Set the dynamic distribution group scope to the OU where the gorup originally resided...' -toscreen

			set-dynamicDistributionGroup -identity $script:newDynamicDLAddress -recipientContainer $script:onpremisesdlConfiguration.organizationalUnit  -domaincontroller $script:adDomainController
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function setOnPremisesDynamicDistributionGroupSettings...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'The properties have been set successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function setOnPremisesDynamicDistributionGroupSettings...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The properties could not be set successfully." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function createAndUpdateMailOnMicrosoftAddress

.DESCRIPTION

This function creates a mail.onmicrosoft.com address and adds it to the group in Office 365.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createAndUpdateMailOnMicrosoftAddress
{
	Param ()

	Begin 
	{
		$functionContactRemoteAddress = $NULL
		$functionEmailAddresses = $NULL	#Utilized to hold working email addresses in the function.

		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function createAndUpdateMailOnMicrosoftAddress...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'THis function creates a mail.onmicrosoft.com address for the group and updates Office 365.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen

		$functionEmailAddresses = $script:newOffice365DLConfiguration.emailAddresses

		#Iterate through all proxy addresses to find the remote routing address.
		#This needs to be removed so that it can be stamped on the mail contact matching this group.
		
		foreach ( $emailAddress in $functionEmailAddresses)
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Iterating through proxy addresses to find onmicrosoft.com address...' -toscreen

			if ( $emailAddress -like "*.onmicrosoft.com" )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'The onmicrosoft.com address has been found...' -toscreen

				$functionContactRemoteAddress=$emailAddress
			}
		}

		#We now have the onmicrosoft.com address stamped on all objects created in the service.
		#Now we need to take this address and convert it into a mail.onmicrosoft.com address.
		#First split at the @ then split at the . so we can inject the mail portion of the address.
		#Then take the entire address and make it lower case.

		$functionContactRemoteAddress = $functionContactRemoteAddress -split "@"

		$functionContactRemoteAddressDomain = $functionContactRemoteAddress[1] -split "\."

		$script:remoteRoutingAddress = $functionContactRemoteAddress[0] + "@" + $functionContactRemoteAddressDomain[0] + ".mail." + $functionContactRemoteAddressDomain[1] + "." + $functionContactRemoteAddressDomain[2]

		$script:remoteRoutingAddress = $script:remoteRoutingAddress.ToLower()
	}
	Process 
	{
		Try 
		{
			set-O365DistributionGroup -identity $script:newOffice365DLConfiguration.Alias -EmailAddresses @{add=$script:remoteRoutingAddress}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function createAndUpdateMailOnMicrosoftAddress...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'The Address was successfully created and updated.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function createAndUpdateMailOnMicrosoftAddress...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The address could not be updated." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function createRemoteRoutingContact

.DESCRIPTION

This function creates a mail enabled contact to allow routing of on premises emails to the new cloud distribuiton group.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function createRemoteRoutingContact
{
	Param ()

	Begin 
	{
		$functionContactRemoteAddress = $NULL

		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function createRemoteRoutingContact...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function creates a mail enabled contact that routes on premises email to the migrated distribution group.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen

		$functionOrganizationalUnit = $NULL

		#Establish contacts random name as previous group name + migratedByScript
		
		$script:randomContactName = $script:onpremisesdlConfiguration.name+"-MigratedByScript"

		#Set the OU to create the contact to match the OU of the original group.

		$functionOrganizationalUnit = $script:onpremisesdlConfiguration.organizationalUnit

		#In the function where we provision the dynamic ditribution group we capture the mail.onmicrosoft.com address.
		#This address should be used on the mail contact as the remote routing address.
		#It is possible that depending on email address policy configuration groups did not get a mail.onmicrosoft.com address when the hybrid configuration wizard was run.
		#This results in the group not having the address in the service.
		#If thhis is the case - we need to create one to ensure the secured connector it utilized cross premises.

		if ( $script:remoteRoutingAddress -eq $NULL )
		{
			createAndUpdateMailOnMicrosoftAddress
		}
	}
	Process 
	{
		Try 
		{
			#Create the mail contact using the previous groups routing address, the random name, and in the original OU.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Creating the mail enabled contact in the original OU as the migrated group...' -toscreen

			new-mailContact -name $script:randomContactName -externalEmailAddress $script:remoteRoutingAddress -organizationalUnit $functionOrganizationalUnit -domaincontroller $script:adDomainController
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function createRemoteRoutingContact...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'The contact was created successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function createRemoteRoutingContact...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The contact was not created successfully." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function setRemoteRoutingContactSettings

.DESCRIPTION

This function sets the settings of the contact for remote routing.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>
Function setRemoteRoutingContactSettings
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function setOnPremisesDynamicDistributionGroupSettings...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This funciton sets the properties of the remote routing contact.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen

		#Get the mail contact created to replace the distribution group.
		
		$functionContact = get-mailcontact -identity $script:randomContactName -domaincontroller $script:adDomainController
		$functionPrimarySMTPAddress = $NULL

		#Iterate through the list of proxy addresses and find the first address that is not an onmicrosoft.com address and that is an SMTP address.

		foreach ( $emailAddress in $functionContact.emailAddresses)
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Iterating through contact email addresses...' -toscreen

			if ( ($emailaddress -like "smtp:*") -and ($emailAddress -notlike "*.onmicrosoft.com") )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message 'Located first non-onmicrosoft.com proxy address...' -toscreen

				$functionPrimarySMTPAddress=$emailAddress
				$functionPrimarySMTPAddress=$functionPrimarySMTPAddress.trimstart("smtp:")
				break
			}
		}
	}
	Process 
	{
		Try 
		{
			#Set the mail contact attributes.
			#Set custom attribute 1 to MigratedBySCript to match the dynamic DL.
			#Set custom attribute 2 to the primary SMTP address of the original group.  We'll use this to match to the group in Office 365 to preserve settings later.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Setting initial properties of the mail contact...' -toscreen

			Set-mailContact -Identity $script:randomContactName -HiddenFromAddressListsEnabled $TRUE -domaincontroller $script:adDomainController -CustomAttribute1 "MigratedByScript" -CustomAttribute2 $script:onpremisesdlConfiguration.primarySMTPAddress
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Disabling automatinc email address policy assignment to the mail contact.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Disabling automatic email address policy on the mail contact...' -toscreen

			Set-mailContact -Identity $script:randomContactName -emailAddressPolicyEnabled:$FALSE -domaincontroller $script:adDomainController
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Setting the primary SMTP address of the mail contac to be any of the non-onmicrosoft.com email addresses -> doesn't matter which one -> causmetic.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Setting primary SMTP address of the mail contact to be any of the valid proxies previously located...' -toscreen

			Set-mailContact -Identity $script:randomContactName -primarySMTPAddress $functionPrimarySMTPAddress  -domaincontroller $script:adDomainController
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Remove the remote routing address from the list of proxy addresses.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Removing the remote routing address from the proxy addresses...' -toscreen

			Set-mailContact -Identity $script:randomContactName -EmailAddresses @{remove=$script:remoteRoutingAddress}  -domaincontroller $script:adDomainController
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Set the master account sid to Self.
			#This is necessary to trick proeprties like ManagedBY which require the contact to have a SID.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Setting the master account SID to self...' -toscreen

			invoke-command -ScriptBlock { Set-ADObject $args[0] -replace @{"msExchMasterAccountSid"=$args[1]} } -ArgumentList $functionContact.distinguishedName,$script:wellKnownSelfAccountSid -Session $script:onPremisesADDomainControllerPowerShellSession
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Record the settings of the new mail contact to a variable for later use.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Recording the new mail contact information to a variable...' -toscreen

			$script:onPremisesNewContactConfiguration = Get-mailContact -identity $script:randomContactName -domaincontroller $script:adDomainController
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function setOnPremisesDynamicDistributionGroupSettings...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'The properties have been set successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function setOnPremisesDynamicDistributionGroupSettings...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The properties could not be set successfully." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function resetDLMemberOf

.DESCRIPTION

This function adds the mail enabled contact back to the groups it was previously a member of as a distribution group.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function resetDLMemberOf
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function resetDLMemberOf...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function adds the mail contact back to the groups it was previously a member of.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen

		$functionContact = get-mailcontact -identity $script:randomContactName -domainController $script:adDomainController
	}
	Process 
	{
		Try 
		{
			foreach ($functionGroup in $script:onPremisesDLMemberOf)
			{
                $loopGroup=$functionGroup.distinguishedName
				$loopContact=$functionContact.distinguishedName
				
				Write-LogInfo -LogPath $script:sLogFile -Message "Adding '$loopContact' to group '$loopGroup'" -toscreen

				invoke-command -ScriptBlock { set-adgroup -identity $args[0] -add @{'member'=$args[1]} } -ArgumentList $loopGroup,$loopContact -Session $script:onPremisesADDomainControllerPowerShellSession
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function resetDLMemberOf...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution group successfully updated.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function resetDLMemberOf...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "Distribution group could not be successfully update." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function resetCloudDLMemberOf

.DESCRIPTION

This function adds the new office 365 distribution groups back to the groups that it was in before deletion.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function resetCloudDLMemberOf
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function resetCloudDLMemberOf...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function adds the new office 365 distribution groups back to the groups that it was in before deletion..' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen

		$functionGroupName = $script:newOffice365DLConfiguration.identity
		[int]$functionCounter = 0
	}
	Process 
	{
		Try 
		{
			foreach ($functionGroup in $script:originalO365MemberOf)
			{
                $loopGroup=$functionGroup.identity
				
				Write-LogInfo -LogPath $script:sLogFile -Message "Adding '$functionGroupName' to group '$loopGroup'" -toscreen

				add-o365distributionGroupMember -identity $loopGroup -member $functionGroupName

				if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
				{
					refreshOffice365PowerShellSession

					$functionCounter=0
				}
				else
				{
					$functionCounter++	
				}
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function resetCloudDLMemberOf...' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Distribution group successfully updated.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message ' ' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function resetCloudDLMemberOf...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "Distribution group could not be successfully update." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function resetOriginalDistributionListSettings

.DESCRIPTION

This function resets settings of distribution lists on premises for a DL that was delete and converted to a contact.

.PARAMETER OperationType

String $arrayName - specifies the script variable array name to work with.

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>
Function resetOriginalDistributionListSettings
{
	Begin 
	{
		[array]$functionArray = @()
		$functionGroup=$NULL

		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function resetOriginalDistributionListSettings...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function takes an array of settings from the deleted DL and resets them...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
	}
	Process 
	{
		#Scan each array from the orignal DL settings and if not null take action.
		
        if ( $script:originalGrantSendOnBehalfTo -ne $NULL ) 
        {
			#The converted group had send as rights to other groups - reset those groups to the mail contact.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing send on behalf to...' -toscreen

			$functionArray = $script:originalGrantSendOnBehalfTo

            foreach ( $member in $functionArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to send on behalf to '$member.primarySMTPAddress -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gathering groups current grant sent on behalf settings... ' -ToScreen
					
						$functionGroup=(get-distributiongroup -identity $member.PrimarySMTPAddress -domainController $script:adDomainController).GrantSendOnBehalfTo
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the mail contact identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding mail contact to send on behalf and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

						$functionGroup+=$script:onPremisesNewContactConfiguration.primarySMTPAddress
						set-distributiongroup -identity $member.PrimarySMTPAddress -GrantSendOnBehalfTo $functionGroup -domainController $script:adDomainController -BypassSecurityGroupManagerCheck
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
            }
		}

        if ( $script:originalAcceptMessagesFrom -ne $NULL  )
        {
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing accept messages on from senders or members...' -toscreen

			#The group had accept only from set on other groups - add mail contact back.

			$functionArray =  $script:originalAcceptMessagesFrom

            foreach ( $member in $functionArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to send on accept messsages only from senders or members ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gathering groups current accept messages from settings... ' -ToScreen

						$functionGroup=(get-distributiongroup -identity $member.PrimarySMTPAddress -domainController $script:adDomainController).AcceptMessagesOnlyFromSendersorMembers  
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the mail contact identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding mail contact to accept messages list and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
						$functionGroup+=$script:onPremisesNewContactConfiguration.primarySMTPAddress
						$functionGroup
						set-distributiongroup -identity $member.PrimarySMTPAddress -AcceptMessagesOnlyFromSendersorMembers $functionGroup -domainController $script:adDomainController -BypassSecurityGroupManagerCheck
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
            }
		}

		if ( $script:originalRejectMessagesFrom -ne $NULL )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing reject messages on from senders or members...' -toscreen

			#The converted group had reject set on other groups - add mail contact.

			$functionArray = $script:originalRejectMessagesFrom

			foreach ( $member in $functionArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to reject messages from senders or members ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gatheing groups current reject from settings... ' -ToScreen

						$functionGroup=(get-distributiongroup -identity $member.PrimarySMTPAddress -domainController $script:adDomainController).RejectMessagesFromSendersOrMembers
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the mail contact identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding mail contact to reject from list and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
						$functionGroup+=$script:onPremisesNewContactConfiguration.primarySMTPAddress
					
						set-distributiongroup -identity $member.PrimarySMTPAddress -RejectMessagesFromSendersOrMembers $functionGroup -domainController $script:adDomainController -BypassSecurityGroupManagerCheck
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
			}
		}

		if ( $script:originalBypassModerationFromSendersOrMembers -ne $NULL )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing bypass messages on from senders or members...' -toscreen

			#Group had bypass moderation from senders or members set on other groups.  Add mail contact to bypass moderation.

			$functionArray = $script:originalBypassModerationFromSendersOrMembers 

			foreach ( $member in $functionArray )
            {
				#Get the distribution list that had the group originall on bypass full bypass list to a variable.

				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to bypass moderation from senders or members ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.PrimarySMTPAddressOrUPN -ToScreen

				if ( $member.primarySMTPAddressorUPN -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gatheing groups current bypass moderation settings... ' -ToScreen

						$functionGroup=(get-distributiongroup -identity $member.PrimarySMTPAddressOrUPN -domainController $script:adDomainController).BypassModerationFromSendersOrMembers  
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the mail contact identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding mail contact to bypass list and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddressorUPN -ToScreen
						$functionGroup+=$script:onPremisesNewContactConfiguration.primarySMTPAddress

						set-distributiongroup -identity $member.PrimarySMTPAddressOrUPN -BypassModerationFromSendersOrMembers $functionGroup -domainController $script:adDomainController -BypassSecurityGroupManagerCheck
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
			}
		}
		if ( $script:originalManagedBy -ne $NULL )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing managed by...' -toscreen

			#Group had bypass moderation from senders or members set on other groups.  Add mail contact to bypass moderation.

			$functionArray = $script:originalManagedBy

			foreach ( $member in $functionArray )
            {
				#Get the distribution list that had the group originall on bypass full bypass list to a variable.

				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to managed by: ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.PrimarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gatheing groups current managed by settings... ' -ToScreen

						$functionGroup=(get-distributiongroup -identity $member.PrimarySMTPAddress -domainController $script:adDomainController).managedBy
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the mail contact identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding mail contact to managed by and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
						$functionGroup+=$script:onPremisesNewContactConfiguration.primarySMTPAddress

						set-distributiongroup -identity $member.PrimarySMTPAddress -ManagedBy $functionGroup -domainController $script:adDomainController -BypassSecurityGroupManagerCheck
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
			}
		}

		if ( $script:originalForwardingAddress -ne $NULL )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing forwading address...' -toscreen

			#Group had forwarding address on other groups.  Add mail contact to forwarding address.

			$functionArray = $script:originalForwardingAddress

			foreach ( $member in $functionArray )
            {
				#Get the distribution list that had the group originall on bypass full bypass list to a variable.

				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding forwarding Address... ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.PrimarySMTPAddress -ToScreen

				Try
				{
					#Set the forwarding address of the mailbox.

					Write-LogInfo -LogPath $script:sLogFile -Message 'Adding forwarding address to the mailbox.... ' -ToScreen
					Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
					
					set-mailbox -identity $member.PrimarySMTPAddress -forwardingAddress $script:onPremisesNewContactConfiguration.identity -domainController $script:adDomainController
				}
				Catch
				{
					Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
					cleanupSessions
					Stop-Log -LogPath $script:sLogFile -ToScreen
					archiveFiles
					Break
				}
			}
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function resetOriginalDistributionListSettings...' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message 'The array was built successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function resetOriginalDistributionListSettings...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The array could not be built successfully - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function resetCloudDistributionListSettings

.DESCRIPTION

This function resets all cloud only DL multi-valued attributes where the migrated DL had settings.

.PARAMETER OperationType

String $arrayName - specifies the script variable array name to work with.

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>
Function resetCloudDistributionListSettings
{
	Begin 
	{
		[array]$functionArray = @()
		$functionGroup=$NULL
		[int]$functionCounter=0

		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function resetCloudDistributionListSettings...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'This function resets all cloud only DL multi-valued attributes where the migrated DL had settings....' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
	}
	Process 
	{
		#Scan each array from the orignal DL settings and if not null take action.
		
        if ( $script:originalO365GrantSendOnBehalfTo -ne $NULL ) 
        {
			#The converted DL had send as rights to other groups - reset those groups to the migrated DL.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing send on behalf to...' -toscreen

			$functionArray = $script:originalO365GrantSendOnBehalfTo

            foreach ( $member in $functionArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to send on behalf to '$member.primarySMTPAddress -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gathering groups current grant sent on behalf settings... ' -ToScreen
					
						$functionGroup=(get-o365Distributiongroup -identity $member.PrimarySMTPAddress).GrantSendOnBehalfTo
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the distribution group identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding distribution group to send on behalf and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

						$functionGroup+=$script:newOffice365DLConfiguration.primarySMTPAddress
						set-o365Distributiongroup -identity $member.PrimarySMTPAddress -GrantSendOnBehalfTo $functionGroup -BypassSecurityGroupManagerCheck

						if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
						{
							refreshOffice365PowerShellSession
						}
						else
						{
							$functionCounter++	
						}
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
            }
		}

		$functionCounter=0

        if ( $script:originalO365AcceptMessagesOnlyFromDLMembers -ne $NULL  )
        {
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing accept messages on from senders or members...' -toscreen

			#The group had accept only from set on other groups - add the group back.

			$functionArray =  $script:originalO365AcceptMessagesOnlyFromDLMembers

            foreach ( $member in $functionArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to send on accept messsages only from senders or members ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gathering groups current accept messages from settings... ' -ToScreen

						$functionGroup=(get-o365Distributiongroup -identity $member.PrimarySMTPAddress).AcceptMessagesOnlyFromSendersorMembers  
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the distribution group identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding group to accept messages list and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
						$functionGroup+=$script:newOffice365DLConfiguration.primarySMTPAddress
						$functionGroup
						set-o365Distributiongroup -identity $member.PrimarySMTPAddress -AcceptMessagesOnlyFromSendersorMembers $functionGroup -BypassSecurityGroupManagerCheck

						if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
						{
							refreshOffice365PowerShellSession
						}
						else
						{
							$functionCounter++	
						}
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
            }
		}

		$functionCounter=0

		if ( $script:originalO365RejectMessagesFromDLMembers -ne $NULL )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing reject messages on from senders or members...' -toscreen

			#The converted group had reject set on other groups - add converted group.

			$functionArray = $script:originalO365RejectMessagesFromDLMembers

			foreach ( $member in $functionArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to reject messages from senders or members ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gatheing groups current reject from settings... ' -ToScreen

						$functionGroup=(get-o365Distributiongroup -identity $member.PrimarySMTPAddress).RejectMessagesFromSendersOrMembers
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the group identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding group to reject from list and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
						$functionGroup+=$script:newOffice365DLConfiguration.primarySMTPAddress
					
						set-o365Distributiongroup -identity $member.PrimarySMTPAddress -RejectMessagesFromSendersOrMembers $functionGroup -BypassSecurityGroupManagerCheck

						if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
						{
							refreshOffice365PowerShellSession
						}
						else
						{
							$functionCounter++	
						}
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
			}
		}
		
		$functionCounter=0

		if ( $script:originalO365BypassModerationFromSendersOrMembers -ne $NULL )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing bypass messages on from senders or members...' -toscreen

			#Group had bypass moderation from senders or members set on other groups.  Add group to bypass moderation.

			$functionArray = $script:originalO365BypassModerationFromSendersOrMembers

			foreach ( $member in $functionArray )
            {
				#Get the distribution list that had the group originall on bypass full bypass list to a variable.

				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to bypass moderation from senders or members ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.PrimarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddressorUPN -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gatheing groups current bypass moderation settings... ' -ToScreen

						$functionGroup=(get-o365Distributiongroup -identity $member.PrimarySMTPAddress).BypassModerationFromSendersOrMembers  
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the group identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding mail contact to bypass list and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
						$functionGroup+=$script:newOffice365DLConfiguration.primarySMTPAddress

						set-o365Distributiongroup -identity $member.PrimarySMTPAddress -BypassModerationFromSendersOrMembers $functionGroup -BypassSecurityGroupManagerCheck

						if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
						{
							refreshOffice365PowerShellSession
						}
						else
						{
							$functionCounter++	
						}
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
			}
		}

		$functionCounter=0

		if ( $script:originalO365ManagedBy -ne $NULL )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing managed by...' -toscreen

			#Group had bypass moderation from senders or members set on other groups.  Add group to bypass moderation.

			$functionArray = $script:originalO365ManagedBy

			foreach ( $member in $functionArray )
            {
				#Get the distribution list that had the group originall on bypass full bypass list to a variable.

				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to managed by: ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.PrimarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gatheing groups current managed by settings... ' -ToScreen

						$functionGroup=(get-o365Distributiongroup -identity $member.PrimarySMTPAddress).managedBy
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the group identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding mail contact to managed by and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
						$functionGroup+=$script:newOffice365DLConfiguration.primarySMTPAddress

						set-o365Distributiongroup -identity $member.PrimarySMTPAddress -ManagedBy $functionGroup -BypassSecurityGroupManagerCheck

						if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
						{
							refreshOffice365PowerShellSession
						}
						else
						{
							$functionCounter++	
						}

					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
			}
		}

		$functionCounter=0

		if ( $script:originalO365ForwardingAddress -ne $NULL )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing forwading address...' -toscreen

			#Group had forwarding address on other groups.  Add group to forwarding address.

			$functionArray = $script:originalO365ForwardingAddress

			foreach ( $member in $functionArray )
            {
				#Get the distribution list that had the group originall on bypass full bypass list to a variable.

				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding forwarding Address... ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.PrimarySMTPAddress -ToScreen

				Try
				{
					#Set the forwarding address of the mailbox.

					Write-LogInfo -LogPath $script:sLogFile -Message 'Adding forwarding address to the mailbox.... ' -ToScreen
					Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
					
					set-o365Mailbox -identity $member.PrimarySMTPAddress -forwardingAddress $script:newOffice365DLConfiguration.identity

					if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
					{
						refreshOffice365PowerShellSession
					}
					else
					{
						$functionCounter++	
					}
				}
				Catch
				{
					Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
					cleanupSessions
					Stop-Log -LogPath $script:sLogFile -ToScreen
					archiveFiles
					Break
				}
			}
		}

		$functionCounter=0

		if ( $script:originalO365GroupAcceptMessagesOnlyFromDLMembers -ne $NULL  )
        {
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing accept messages on from senders or members...' -toscreen

			#The group had accept only from set on other groups - add the group back.

			$functionArray =  $script:originalO365GroupAcceptMessagesOnlyFromDLMembers

            foreach ( $member in $functionArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to send on accept messsages only from senders or members ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gathering groups current accept messages from settings... ' -ToScreen

						$functionGroup=(get-o365UnifiedGroup -identity $member.PrimarySMTPAddress).AcceptMessagesOnlyFromSendersorMembers  
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the distribution group identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding group to accept messages list and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
						$functionGroup+=$script:newOffice365DLConfiguration.primarySMTPAddress
						$functionGroup
						set-o365Unifiedgroup -identity $member.PrimarySMTPAddress -AcceptMessagesOnlyFromSendersorMembers $functionGroup

						if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
						{
							refreshOffice365PowerShellSession
						}
						else
						{
							$functionCounter++	
						}

					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
            }
		}

		$functionCounter=0

		if ( $script:originalO365GroupRejectMessagesFromDLMembers -ne $NULL )
		{
			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing reject messages on from senders or members...' -toscreen

			#The converted group had reject set on other groups - add converted group.

			$functionArray = $script:originalO365GroupRejectMessagesFromDLMembers

			foreach ( $member in $functionArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to reject messages from senders or members ' -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gatheing groups current reject from settings... ' -ToScreen

						$functionGroup=(get-o365UnifiedGroup -identity $member.PrimarySMTPAddress).RejectMessagesFromSendersOrMembers
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the group identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding group to reject from list and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen
						$functionGroup+=$script:newOffice365DLConfiguration.primarySMTPAddress
					
						set-o365Unifiedgroup -identity $member.PrimarySMTPAddress -RejectMessagesFromSendersOrMembers $functionGroup

						if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
						{
							refreshOffice365PowerShellSession
						}
						else
						{
							$functionCounter++	
						}
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
			}
		}

		$functionCounter=0

		if ( $script:originalO365GroupGrantSendOnBehalfTo -ne $NULL ) 
        {
			#The converted DL had send as rights to other groups - reset those groups to the migrated DL.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Processing send on behalf to...' -toscreen

			$functionArray = $script:originalO365GroupGrantSendOnBehalfTo

            foreach ( $member in $functionArray )
            {
				Write-LogInfo -LogPath $script:sLogFile -Message 'Adding to send on behalf to '$member.primarySMTPAddress -toscreen
				Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

				if ( $member.primarySMTPAddress -ne $script:onpremisesdlConfiguration.primarySMTPAddress )
				{
					Try
					{
						Write-LogInfo -LogPath $script:sLogFile -Message 'Gathering groups current grant sent on behalf settings... ' -ToScreen
					
						$functionGroup=(get-o365UnifiedGroup -identity $member.PrimarySMTPAddress).GrantSendOnBehalfTo
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
					Try
					{
						#Add the distribution group identity to the list and then restamp the entire list.
						#This was done because array operations @{ADD=*} did not work against this attribute.

						Write-LogInfo -LogPath $script:sLogFile -Message 'Adding distribution group to send on behalf and stamping full list on group... ' -ToScreen
						Write-LogInfo -LogPath $script:sLogFile $member.primarySMTPAddress -ToScreen

						$functionGroup+=$script:newOffice365DLConfiguration.primarySMTPAddress
						set-o365Unifiedgroup -identity $member.PrimarySMTPAddress -GrantSendOnBehalfTo $functionGroup

						if ( $functionCounter -gt $script:refreshPowerShellSessionCounter )
						{
							refreshOffice365PowerShellSession
						}
						else
						{
							$functionCounter++	
						}
					}
					Catch
					{
						Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
						cleanupSessions
						Stop-Log -LogPath $script:sLogFile -ToScreen
						archiveFiles
						Break
					}
				}
            }
		}
	}
	End 
	{
		If ($?) 
		{
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message 'Exiting function resetCloudDistributionListSettings...' -toscreen
            Write-LogInfo -LogPath $script:sLogFile -Message 'The array was built successfully.' -toscreen
			Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		}
		else
		{
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message 'Exiting function resetCloudDistributionListSettings...' -toscreen
			Write-LogError -LogPath $script:sLogFile -Message "The array could not be built successfully - exiting." -toscreen
			Write-LogError -LogPath $script:sLogFile -Message $error[0] -toscreen
			Write-LogError -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
		}
	}
}

<#
*******************************************************************************************************

Function recordOriginalO365MultivaluedAttributes

.DESCRIPTION

Records information regarding the multi-valued attributes of Office 365 cloud only distribution lists for preservation.

.PARAMETER 

NONE

.INPUTS

NONE

.OUTPUTS 

NONE

*******************************************************************************************************
#>

Function recordOriginalO365MultivaluedAttributes
{
	Param ()

	Begin 
	{
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Entering function recordOriginalO365MultivaluedAttributes...' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message 'Records information regarding the multi-valued attributes of Office 365 cloud only distribution lists for preservation.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message '******************************************************************' -toscreen

		#Record the identity of the moved distribution list (updated post move so we have correct identity)
		
		$functionGroupIdentity = $script:office365DLConfiguration.identity.tostring()	#Function variable to hold the identity of the group.
		$functionGroupForwardingIdentity = ConvertFrom-DN -DistinguishedName $script:office365DLConfiguration.distinguishedName
		$functionFixedGroupForwardingIdentity = $($functionGroupForwardingIdentity.replace("'","''"))
		$functionCommand = $NULL	#Holds the expression that we will be executing to determine multi-valued membership.
		$functionRecipientObject = $NULL
		[array]$functionAllCloudOnlyGroups = $NULL
		[array]$functionAllOffice365Groups = $NULL
		
		Write-LogInfo -LogPath $script:sLogFile -Message 'The following group identity is the filtered name.' -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $functionGroupIdentity -toscreen
	}
	Process 
	{
		Try 
		{
			#Getting all distribution groups where they are cloud only.

			#Filters in office 365 are not necessarily reliable based on testing.
			#What we will do here is gather all cloud only group objects into a variable for the purposes of working through it.
			#This should limit the query to the service against this call - and from there we will filter out all the things we want.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gathering all cloud only distribution lists...' -toscreen

			$functionCommand = "get-o365DistributionGroup -resultsize unlimited -filter { IsDirSynced -eq `$FALSE}"

			$functionAllCloudOnlyGroups = Invoke-Expression $functionCommand
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Filters in office 365 are not necessarily reliable based on testing.
			#What we will do here is gather all cloud only group objects into a variable for the purposes of working through it.
			#This should limit the query to the service against this call - and from there we will filter out all the things we want.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gathering all Office 365 Groups...' -toscreen

			$functionAllOffice365Groups = get-O365UnifiedGroup -resultsize unlimited
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Iterate through all groups locating those that have grant send on behalf to set to the identity.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all grantSendOnBehalfTo for the identity...' -toscreen

			foreach ( $functionGroup in $functionAllCloudOnlyGroups )
			{
				if ( $functionGroup.grantSendOnBehalfTo -eq $functionFixedGroupForwardingIdentity )
				{
					Write-LogInfo -LogPath $script:sLogFile -Message $functionGroup.primarySMTPAddress -ToScreen
					$script:originalO365GrantSendOnBehalfTo+=$functionGroup
				}
				if ( $functionGroup.AcceptMessagesOnlyFromDLMembers -eq $functionFixedGroupForwardingIdentity )
				{
					Write-LogInfo -LogPath $script:sLogFile -Message $functionGroup.primarySMTPAddress -ToScreen
					$script:originalO365AcceptMessagesOnlyFromDLMembers+=$functionGroup
				}
				if ( $functionGroup.ManagedBy -eq $functionFixedGroupForwardingIdentity )
				{
					Write-LogInfo -LogPath $script:sLogFile -Message $functionGroup.primarySMTPAddress -ToScreen
					$script:originalO365ManagedBy+=$functionGroup
				}
				if ( $functionGroup.RejectMessagesFromDLMembers -eq $functionFixedGroupForwardingIdentity )
				{
					Write-LogInfo -LogPath $script:sLogFile -Message $functionGroup.primarySMTPAddress -ToScreen
					$script:originalO365RejectMessagesFromDLMembers+=$functionGroup
				}
				if ( $functionGroup.bypassModerationFromSendersOrMembers -eq $functionFixedGroupForwardingIdentity )
				{
					Write-LogInfo -LogPath $script:sLogFile -Message $functionGroup.primarySMTPAddress -ToScreen
					$script:originalO365BypassModerationFromSendersOrMembers+=$functionGroup
				}
			}	
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Interate through all mailboxes and determine any that have a forwarding address of the DL to be migrated.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all cloud mailboxes enabled for forwarding to the migrated DL for the identity...' -toscreen
			
			$functionCommand = "get-o365Mailbox -resultsize unlimited -Filter { ForwardingAddress -eq '$functionFixedGroupForwardingIdentity' }"

			$script:originalO365ForwardingAddress = invoke-expression $functionCommand

			foreach ( $member in $script:originalO365ForwardingAddress )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message $member.primarySMTPAddress -ToScreen
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Iterate through all groups and see if we can find any cloud only groups that this user is a member of.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all cloud only distribution groups that this user is a member of...' -toscreen
			
			$functionCommand = "get-o365Recipient -resultsize unlimited -Filter { Members -eq '$functionFixedGroupForwardingIdentity' }"

			$script:originalO365MemberOf = invoke-expression $functionCommand

			foreach ( $member in $script:originalO365MemberOf )
			{
				Write-LogInfo -LogPath $script:sLogFile -Message $member.primarySMTPAddress -ToScreen
			}
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Iterate through all grant send on behalf to in Office 365 groups.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all grantSendOnBehalfTo for the identity...' -toscreen

			foreach ( $functionGroup in $functionAllOffice365Groups )
			{
				if ( $functionGroup.grantSendOnBehalfTo -eq $functionFixedGroupForwardingIdentity )
				{
					Write-LogInfo -LogPath $script:sLogFile -Message $functionGroup.primarySMTPAddress -ToScreen
					$script:originalO365GroupGrantSendOnBehalfTo+=$functionGroup
				}
			}	
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Iterate through all office 365 groups locating those that have grant send on behalf to set to the identity.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all acceptMessageOnlyFromDLMembers for the identity...' -toscreen

			foreach ( $functionGroup in $functionAllOffice365Groups )
			{
				if ( $functionGroup.AcceptMessagesOnlyFromDLMembers -eq $functionFixedGroupForwardingIdentity )
				{
					Write-LogInfo -LogPath $script:sLogFile -Message $functionGroup.primarySMTPAddress -ToScreen
					$script:originalO365GroupAcceptMessagesOnlyFromDLMembers+=$functionGroup
				}
			}	
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
		Try 
		{
			#Iterate through all office 365 groups locating those that have grant send on behalf to set to the identity.

			Write-LogInfo -LogPath $script:sLogFile -Message 'Gather all rejectMessagesFromDLMembers for the identity...' -toscreen

			foreach ( $functionGroup in $functionAllOffice365Groups )
			{
				if ( $functionGroup.RejectMessagesFromDLMembers -eq $functionFixedGroupForwardingIdentity )
				{
					Write-LogInfo -LogPath $script:sLogFile -Message $functionGroup.primarySMTPAddress -ToScreen
					$script:originalO365GroupRejectMessagesFromDLMembers+=$functionGroup
				}
			}	
		}
		Catch 
		{
			Write-LogError -LogPath $script:sLogFile -Message $_.Exception -toscreen
			cleanupSessions
			Stop-Log -LogPath $script:sLogFile -ToScreen
			archiveFiles
			Break
		}
	}
}


#-----------------------------------------------------------[Execution]------------------------------------------------------------

#Create log file for operations within this script.

New-Item -ItemType Directory -Path $script:sLogPath -Force

Start-Log -LogPath $script:sLogPath -LogName $script:sLogName -ScriptVersion $script:sScriptVersion -ToScreen

recordCommandLineOptions #Records the command line options specified in the command call.

establishOnPremisesCredentials  #Function call to import and populate on premises credentials.

establishOffice365Credentials  #Function call to import and populate Office 365 credentials.

#If interactive credentials were requested - validate that nothing was passsed as null.

If ( $requireInteractiveCredentials -eq $TRUE )
{
	validateInteractiveCredentials ( $script:onPremisesCredential ) ( "OnPremises" )

	validateInteractiveCredentials ( $script:office365Credential  ) ( "Office365" )
}

createOnPremisesPowershellSession  #Creates the on premises powershell session to Exchange.

createOffice365PowershellSession  #Creates the Office 365 powershell session.

createOnPremisesAADConnectPowershellSession  #Create the on premises AAD Connect powershell session.

createOnPremisesADDomainControllerPowershellSession  #Create the on premises AD Domain controller powershell session.

importOnPremisesPowershellSession  #Function call to import the on premises powershell session.

importOffice365PowershellSession  #Function call to import the Office 365 powershell session.

collectOnPremsiesDLConfiguration  #Function call to gather the on premises DL information.

collectOffice365DLConfiguation  #Function call to gather the Office 365 DL information.

performOffice365SafetyCheck #Checks to see if the distribution list provided has already been migrated.

backuponpremisesdlConfiguration  #Function call to write the on premises DL information to XML.

backupOffice365DLConfiguration  #Function call to write the Office 365 DL information to XML.

collectonpremisesdlconfigurationMembership  #Function collects the membership of the on premise DL.

backupOnPremisesdlMembership #Writes the on premises DL membership to XML for protection and auditing.

#Begin processing the on premises membership and multi-valued array.
#We take the array and call build array to create a normalized list of SMTP addresses for later operations.
#This is required since many of the multi-valued attributes store their references as disinguished names - which will not translate to Office 365.
#We then test to ensure that the target recipient exists in office 365 as a mail enabled object or a user.
#We then test to see if the recipient is a group has the group already been migrated.  Groups that are members must be migrated first.
#Several set functions have a counter routine set to the defined value.  When we hit this value - we refresh powershell to office 365.

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing a DL membership array." -ToScreen

if ( $script:onpremisesdlconfigurationMembership -ne $NULL )
{
    buildMembershipArray ( "DLMembership" ) ( "onpremisesdlconfigurationMembershipArray") ($ignoreInvalidDLMember) #This function builds an array of members for the DL <or> multivalued attributes.
    
	if ( $script:onpremisesdlconfigurationMembership.count -gt $script:refreshPowerShellSessionCounter )
	{
		refreshOffice365PowerShellSession #Refreshing the session here since building the membership array can take a while depending on array size.
	}
	
	$script:forCounter=0
	$script:arrayCounter=0
	
	foreach ($member in $script:onpremisesdlconfigurationMembershipArray)
	{
        if ($script:forCounter -gt $script:refreshCounter)
        {
            refreshOffice365PowerShellSession
            $script:forCounter = 0
        }

		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		#It was discovered in a customers environment that the distribution list commands set and add aren't near as specific as other Exchange commands.
		#This means that when you have two users whos SMTP addresses were very similar - you would often get errors during the conversation that adding a member of setting a DL property found multiple users.
		#We added an array counter to each function that counts where were at in processing the member array.
		#Since the test function gets the recipient from office 365 - we now capture the GUID of the object.  This is added to the GUID section of the object created before.
		#We will modify moving forward to add members and set attributes via object GUID - instead of the normalized SMTP address.
		#Use of script based variables here was kinda hokie - but it works.

		$script:onpremisesdlconfigurationMembershipArray[$script:arrayCounter].GUID=$script:arrayGUID

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
        }
        
		$script:forCounter+=1
		$script:arrayCounter+=1
    }
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing a ManagedBy array." -ToScreen

if ( $script:onpremisesdlConfiguration.ManagedBy -ne $NULL )
{
    buildMembershipArray ( "ManagedBy" ) ( "onpremisesdlconfigurationManagedByArray" ) ( $ignoreInvalidManagedByMember )
    
    if ( $script:onpremisesdlConfiguration.ManagedBy.count -gt $script:refreshPowerShellSessionCounter )
	{
		refreshOffice365PowerShellSession #Refreshing the session here since building the membership array can take a while depending on array size.
	}

	$script:arrayCounter=0

	foreach ($member in $script:onpremisesdlconfigurationManagedByArray)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
		}

		$script:onpremisesdlconfigurationManagedByArray[$script:arrayCounter].GUID = $script:arrayGUID

		$script:arrayCounter+=1
	}
}

Write-LogInfo -logpath $script:sLogFile -Message "Begin processing a ModeratedBy array." -ToScreen

if ( $script:onpremisesdlConfiguration.ModeratedBy -ne $NULL )
{
    buildmembershipArray ( "ModeratedBy" ) ( "onpremisesdlconfigurationModeratedByArray" ) 
    
    if ( $script:onpremisesdlConfiguration.ModeratedBy.count -gt $script:refreshPowerShellSessionCounter )
	{
		refreshOffice365PowerShellSession #Refreshing the session here since building the membership array can take a while depending on array size.O
	}

	$script:arrayCounter=0

	foreach ($member in $script:onpremisesdlconfigurationModeratedByArray)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)
		
		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
		}

		$script:onpremisesdlconfigurationModeratedByArray[$script:arrayCounter].GUID = $script:arrayGUID

		$script:arrayCounter+=1
	}
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing a GrantSendOnBehalfTo array" -ToScreen

if ( $script:onpremisesdlConfiguration.GrantSendOnBehalfTo -ne $NULL )
{
    buildmembershipArray ( "GrantSendOnBehalfTo" ) ( "onpremisesdlconfigurationGrantSendOnBehalfTOArray" )
    
    if ( $script:onpremisesdlConfiguration.GrantSendOnBehalfTo.count -gt $script:refreshPowerShellSessionCounter )
	{
		refreshOffice365PowerShellSession #Refreshing the session here since building the membership array can take a while depending on array size.
	}

	$script:arrayCounter=0

	foreach ($member in $script:onpremisesdlconfigurationGrantSendOnBehalfTOArray)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			if ( $script:onpremisesdlConfiguration.primarySMTPAddress -ne $member.PrimarySMTPAddressOrUPN )
			{
				testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
			}
		}

		$script:onpremisesdlconfigurationGrantSendOnBehalfTOArray[$script:arrayCounter].GUID = $script:arrayGUID

		$script:arrayCounter+=1
	}
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing a AcceptMessagesOnlyFromSendersOrMembers array" -ToScreen

if ( $script:onpremisesdlConfiguration.AcceptMessagesOnlyFromSendersOrMembers -ne $NULL )
{
    buildMembershipArray ( "AcceptMessagesOnlyFromSendersOrMembers" ) ( "onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers" )
    
    if ( $script:onpremisesdlConfiguration.AcceptMessagesOnlyFromSendersOrMembers.count -gt $script:refreshPowerShellSessionCounter )
	{
		refreshOffice365PowerShellSession #Refreshing the session here since building the membership array can take a while depending on array size.
	}

	$script:arrayCounter=0

	foreach ($member in $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			if ( $script:onpremisesdlConfiguration.primarySMTPAddress -ne $member.PrimarySMTPAddressOrUPN )
			{
				testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
			}
		}

		$script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers[$script:arrayCounter].GUID = $script:arrayGUID

		$script:arrayCounter+=1
	}
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing RejectMessagesFromSendersOrMembers array" -ToScreen

if ( $script:onpremisesdlConfiguration.RejectMessagesFromSendersOrMembers -ne $NULL)
{
    buildMembershipArray ( "RejectMessagesFromSendersOrMembers" ) ( "onpremisesdlconfigurationRejectMessagesFromSendersOrMembers" )
    
    if ( $script:onpremisesdlConfiguration.RejectMessagesFromSendersOrMembers.count -gt $script:refreshPowerShellSessionCounter )
	{
		refreshOffice365PowerShellSession #Refreshing the session here since building the membership array can take a while depending on array size.
	}

	$script:arrayCounter=0

	foreach ($member in $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			if ( $script:onpremisesdlConfiguration.primarySMTPAddress -ne $member.PrimarySMTPAddressOrUPN )
			{
				testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
			}
		}

		$script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers[$script:arrayCounter].GUID = $script:arrayGUID

		$script:arrayCounter+=1
	}
}

Write-LogInfo -LogPath $script:sLogFile -Message "Begin processing BypassModerationFromSendersOrMembers array" -ToScreen

if ( $script:onpremisesdlConfiguration.BypassModerationFromSendersOrMembers -ne $NULL)
{
    buildMembershipArray ( "BypassModerationFromSendersOrMembers") ( "onPremsiesDLBypassModerationFromSendersOrMembers" )
    
    if ( $script:onpremisesdlConfiguration.BypassModerationFromSendersOrMembership.count -gt $script:refreshPowerShellSessionCounter )
	{
		refreshOffice365PowerShellSession #Refreshing the session here since building the membership array can take a while depending on array size.
	}

	$script:arrayCounter=0

	foreach ($member in $script:onPremsiesDLBypassModerationFromSendersOrMembers)
	{
		testOffice365Recipient ($member.PrimarySMTPAddressOrUPN) ($member.RecipientorUser)

		if ( ( $member.recipientType -eq "MailUniversalSecurityGroup" ) -or ($member.recipientType -eq "MailUniversalDistributionGroup") )
		{
			if ( $script:onpremisesdlConfiguration.primarySMTPAddress -ne $member.PrimarySMTPAddressOrUPN )
			{
				testOffice365GroupMigrated ($member.PrimarySMTPAddressOrUPN)
			}
		}

		$script:onPremsiesDLBypassModerationFromSendersOrMembers[$script:arrayCounter].GUID = $script:arrayGUID

		$script:arrayCounter+=1
	}
}

backupOnPremisesDLArrays

#If the administrator has choosen to preserve the cloud only distribution  group settings - we need to determine everything now. 
#This requires the convert to contact option to be selected.

if ( $convertToContact -eq $TRUE )
{
	if ( $retainO365CloudOnlySettings -eq $TRUE )
	{
		#We need to gather all the settings that are on the existing dir synced distribution group that may be set on cloud only distribution groups.

		recordOriginalO365MultivaluedAttributes

		backupO365RetainedSettings
	}	
}

moveGroupToOU  #Move the group to a non-sync OU to preserve it.

#Replicate each domain controller in the domain.

replicateDomainControllers

#Start countdown for the period of time specified by the variable for post domain controller replication.

Start-PSCountdown -Minutes $script:adDomainReplicationTime -Title "Waiting for domain controller replication" -Message "Waiting for domain controller replication"

Write-LogInfo -LogPath $script:sLogFile -Message "Invoking AADConnect Delta Sync Remotely" -ToScreen

#Invoke ad sync.
#If a sync is already in progress this will return an retryable error condition.
#Set the retry variable to true and loop back through.
#The script enforces that at least one delta sync was triggered by it to ensure that it was run post DL movement to non-sync OU.

do
{
	if ( $script:aadconnectRetryRequired -eq $TRUE )
	{
		Start-PSCountdown -Minutes $script:adDomainReplicationTime -Title "Waiting for previous sync to finishn <or> allowing time for invoked sync to run" -Message "Waiting for previous sync to to finishn <or> allowing time for invoked sync to run"
	}
	invokeADConnect	
	$script:aadconnectRetryRequired = $TRUE	
} while ( $error -ne $NULL )

refreshOffice365PowerShellSession

$error.clear()

#Test to wait until the DL is removed from the service.
#This code continues to try to find the original DL in office 365.
#If the DL is found the retry variable will tirgger us to loop back around and try again.
#When the error condition is encountered the DL is no longer there - good - we can move on.

Write-LogInfo -LogPath $script:sLogFile -Message "Waiting for original DL deletion from Office 365" -ToScreen

do
{
	if ( $script:dlDeletionRetryRequired -eq $TRUE)
	{
		Write-LogInfo -LogPath $script:sLogFile -Message "Waiting for original DL deletion from Office 365" -ToScreen
		Start-PSCountdown -Minutes $script:dlDeletionTime -Title "Waiting for DL deletion to process in Office 365" -Message "Waiting for DL deletion to process in Office 365"
		$error.clear()
	}
	$script:dlDeletionRetryRequired = $TRUE
	$scriptTest=get-o365Recipient -identity $script:onpremisesdlConfiguration.primarySMTPAddress
	
} until ( $error -ne $NULL )

refreshOffice365PowerShellSession

$error.clear()

start-sleep -s 30

createOffice365DistributionList

#Set the settings of the distrbution list.
#For multivalued attributes that are not NULL set the individual multivalued attribute.
#For multivalued attributes trigger the appropriate add function with the operation name and the recipient to add.

Start-Sleep -s 30

setOffice365DistributionListSettings

if ( $script:onpremisesdlconfigurationMembershipArray -ne $NULL)
{
    $script:forCounter=0

	foreach ($member in $script:onpremisesdlconfigurationMembershipArray)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing DL Membership member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.GUID -toscreen
        if ($script:forCounter -gt $script:refreshPowerShellSessionCounter)
        {
            refreshOffice365PowerShellSession
            $script:forCounter=0
        }
        setOffice365DistributionlistMultivaluedAttributes ( "DLMembership" ) ( $member.GUID )
        $script:forCounter+=1
	}
}

if ( $script:onpremisesdlconfigurationManagedByArray -ne $NULL)
{
	foreach ($member in $script:onpremisesdlconfigurationManagedByArray )
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Bypass Managed By member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.GUID -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "ManagedBy" ) ( $member.GUID )
	}
}

if ( $script:onpremisesdlconfigurationModeratedByArray -ne $NULL)
{
	foreach ($member in $script:onpremisesdlconfigurationModeratedByArray)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Moderated By member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.GUID -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "ModeratedBy" ) ( $member.GUID  )
	}
}

if ( $script:onpremisesdlconfigurationGrantSendOnBehalfTOArray -ne $NULL )
{
	foreach ($member in $script:onpremisesdlconfigurationGrantSendOnBehalfTOArray)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Grant Send On Behalf To Array member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.GUID -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "GrantSendOnBehalfTo" ) ( $member.GUID  )
	}
}

if ( $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers -ne $NULL )
{
	foreach ($member in $script:onpremisesdlconfigurationAcceptMessagesOnlyFromSendersOrMembers)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Accept Messages Only From Senders Or Members member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.GUID -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "AcceptMessagesOnlyFromSendersOrMembers" ) ( $member.GUID  )
	}
}

if ( $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers -ne $null)
{
	foreach ($member in $script:onpremisesdlconfigurationRejectMessagesFromSendersOrMembers)
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Reject Messages From Senders Or Members member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.GUID -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "RejectMessagesFromSendersOrMembers" ) ( $member.GUID  )
	}
}

if ( $script:onPremsiesDLBypassModerationFromSendersOrMembers -ne $NULL )
{
	foreach ($member in $script:onPremsiesDLBypassModerationFromSendersOrMembers )
	{
		Write-Loginfo -LogPath $script:sLogFile -Message "Processing Bypass Moderation From Senders Or Members member to Office 365..." -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.PrimarySMTPAddressOrUPN -toscreen
		Write-LogInfo -LogPath $script:sLogFile -Message $member.GUID -toscreen
		setOffice365DistributionlistMultivaluedAttributes ( "BypassModerationFromSendersOrMembers" ) ( $member.GUID  )
	}
}

refreshOffice365PowerShellSession

collectNewOffice365DLInformation  #Collect the new office 365 dl configuration.

collectNewOffice365DLMemberInformation #Collect the new office 365 dl membership information.

backupNewOffice365DLConfiguration  #Backup the office 365 DL configuration.

#Assuming there were members in the DL (why would you migrate otherwise but we'll check anyway...write those members to XML.)

if ($script:newOffice365DLConfigurationMembership -ne $NULL)
{
	backupNewOffice365DLConfigurationMembership
}

if ($convertToContact -eq $TRUE)
{
	#To determine if a group is set on the properties of another groups attributes - we need the group id.  The ID needs to be updated since the groups OU was moved.

	recordMovedOriginalDistributionGroupProperties

	#The administrator has decided the way to preserve as many on premises properties as available.
	#This may add significant time to script execution.
	
	if ( $retainOnPremisesSettings -eq $TRUE )
	{
		#Record the membership of the distribution group in other groups.  This will be utilized to reset the mail contact.

		recordDistributionGroupMembership

		#Write the on premises member of information to XML in case of conversion failure.

		backupOnPremisesMemberOf

		#The distribution list can get set to serveral properties on other lists.
		#The goal of this function is to locate those and record them.
		#If the group migrating has permissions to itself - skip the recoridng as it's not required.

		recordOriginalMultivaluedAttributes

		#Write the multi valued attributes to XML in case of conversion failure.

		backupOnPremisesMultiValuedAttributes
	}
	
	#Remove the on prmeises distribution list that was converted.

	removeOnPremisesDistributionGroup

	#We wil utilize a dynamic distribution group to represent the original group in the GAL.
	#This ensures that under no circumstances can we have an address collission.

	createOnPremisesDynamicDistributionGroup

	#Set the attributes of the created dynamic DL.

	setOnPremisesDynamicDistributionGroupSettings

	#Create the mail enabled contact that the dynamic distribution group will referece to move mail to office 365.

	createRemoteRoutingContact

	setRemoteRoutingContactSettings

	if ( $retainOnPremisesSettings -eq $TRUE )
	{
		resetDLMemberOf
	
		#It is possible that the distribution list has permissions to itself.  The find logic goes through and attempts to locate it - and will find it with permissions to itself.
		#Since we're deleting it it cannot be reset.  Skip this function.
	
		resetOriginalDistributionListSettings
	}

	if ( $retainO365CloudOnlySettings -eq $True )
	{
		resetCloudDLMemberOf

		resetCloudDistributionListSettings
	}
	
	#Replicate each domain controller in the domain.

	replicateDomainControllers

	#Start countdown for the period of time specified by the variable for post domain controller replication.

	Start-PSCountdown -Minutes $script:adDomainReplicationTime -Title "Waiting for domain controller replication" -Message "Waiting for domain controller replication"

	do
	{
		if ( $script:aadconnectRetryRequired -eq $TRUE )
		{
			Start-PSCountdown -Minutes $script:adDomainReplicationTime -Title "Waiting for previous sync to finishn <or> allowing time for invoked sync to run" -Message "Waiting for previous sync to to finishn <or> allowing time for invoked sync to run"
		}
		invokeADConnect	
		$script:aadconnectRetryRequired = $TRUE	
	} while ( $error -ne $NULL )

	$error.clear()
}

cleanupSessions  #Clean up - were outta here.

archiveFiles	#Achive the move files so we have them for future reference.
