##############################################################
##############################################################
#                                                            #
#                                                            #
#              AD Information Gathering Script               #
#                                                            #
#                                                            #
##############################################################
##############################################################
#                                                            #
#                   ****INFORMATION****                      #
#                                                            #
#       This script will gather information about the        #
#      Active Directory environment that you are currently   #
#       in and will output the Forest Name, Domain Name,     #
#      current number of OU's, current number of users,      #
#       current number of groups, and a list of other        #
#         informational items.                               #
#                                                            #
##############################################################
#
#
#
#
#      ***** PRERQUISITE - CHECKING OUTPUT PATH *****
$ExportPath = "C:\temp\AD_Info_Gather_Docs"
If(!(test-path $ExportPath))
{
      New-Item -ItemType Directory -Force -Path $ExportPath
}

 

#  ***** PART ONE - GENERIC AD INFORMATION WITH COUNTS *****

 

Import-Module ActiveDirectory

 

 

$ForestFunction = (Get-ADForest).ForestMode
$DomainFunction = (Get-ADDomain).DomainMode
$SchemaVersion = Get-ADObject (Get-ADRootDSE).schemaNamingContext -Propert objectVersion | Select objectVersion
$FSMOInfra = Get-ADDomain | Select-Object InfrastructureMaster
$FSMOPDC = Get-ADDomain | Select-Object PDCEmulator
$FSMORID = Get-ADDomain | Select-Object RIDMaster
$FSMODOM = Get-ADForest | Select-Object DomainNamingMaster
$FSMOSchema = Get-ADForest | Select-Object SchemaMaster
$ADUserCount = (get-aduser -filter *).count
$ADComputerCount = (get-adcomputer -filter *).count
$ADTotalGroupCount = (get-adgroup -filter *).count
$ADSecGroupCount = (get-adgroup -filter 'GroupCategory -eq "security"').count
$ADDistGroupCount = (get-adgroup -filter 'GroupCategory -eq "distribution"').count
$ADDisableUserCount = (get-aduser -filter * | where {$_.enabled -ne "False"}).count
$ADDisabledComputerCount = (get-adcomputer -Filter * | where {$_.enabled -ne "False"}).count
$ADPrinterCount = (Get-AdObject -filter "objectCategory -eq 'printqueue'").count
$ADContactCount = (Get-ADObject -ldapFilter "(objectclass=contact)").count
$ADOUCount = (Get-ADOrganizationalUnit -Filter *).count

 

$Report1 = @(
    [pscustomobject]@{

 

        ForestFunctionalLevel = $ForestFunction
        DomainFunctionallLevel = $DomainFunction
        SchemaVersion = $SchemaVersion
        FSMOInfrastructure = $FSMOInfra
        FSMOPDCEmulator = $FSMOPDC
        FSMORIDMaster = $FSMORID
        FSMODomainNamingMaster = $FSMODOM
        FSMOSchemaMaster = $FSMOSchema
        UserCount = $ADUserCount
        ComputerCount = $ADComputerCount
        DisabledUseCount = $ADDisableUserCount
        DisabledComputerAccount = $ADDisabledComputerCount
        TotalGroupCount = $ADTotalGroupCount
        SecurityGroupCount = $ADSecGroupCount
        DistributionGroupCount = $ADDistGroupCount
        PrinterCount = $ADPrinterCount
        ContactCount = $ADContactCount
        OUCount = $ADOUCount

 

        }
    )
$Report1 | Export-csv -Path $ExportPath\AD_Generic-Count_Info.csv -NoTypeInformation
