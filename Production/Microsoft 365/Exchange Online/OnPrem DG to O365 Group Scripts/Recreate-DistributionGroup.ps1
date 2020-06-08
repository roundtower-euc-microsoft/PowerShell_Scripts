######################################################################################################
#                                                                                                    #
# Name:        Recreate-DistributionGroup.ps1                                                        #
#                                                                                                    #
# Version:     1.0                                                                                   #
#                                                                                                    #
# Description: Copies attributes of a synchronized group to a placeholder group and CSV file.  After #
#              initial export of group attributes, the on-premises group can have the attribute      #
#              "AdminDescription" set to "Group_NoSync" which will stop it from be synchronized.     #
#              The "-Finalize" switch can then be used to write the addresses to the new group and   #
#              convert the name.  The final group will be a cloud group with the same attributes as  #
#              the previous but with the additional ability of being able to be "self-managed".      #
#              Once the contents of the new group are validated, the on-premises group can be        #
#              deleted.                                                                              #
#                                                                                                    #
# Requires:    Remote PowerShell Connection to Exchange Online                                       #
#                                                                                                    #
# Author:      Joe Palarchio                                                                         #
#                                                                                                    #
# Usage:       Additional information on the usage of this script can found at the following         #
#              blog post:  http://blogs.perficient.com/microsoft/?p=32092                            #
#                                                                                                    #
# Disclaimer:  This script is provided AS IS without any support. Please test in a lab environment   #
#              prior to production use.                                                              #
#                                                                                                    #
######################################################################################################


<#
	.PARAMETER Group
		Name of group to recreate.

	.PARAMETER CreatePlaceHolder
		Create placeholder group.

	.PARAMETER Finalize
		Convert placeholder group to final group.

    	.EXAMPLE #1
        	.\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -CreatePlaceHolder

    	.EXAMPLE #2
        	.\Recreate-DistributionGroup.ps1 -Group "DL-Marketing" -Finalize
#>


Param(
    [Parameter(Mandatory=$True)]
        [string]$Group,
    [Parameter(Mandatory=$False)]
        [switch]$CreatePlaceHolder,
    [Parameter(Mandatory=$False)]
        [switch]$Finalize
)

$ExportDirectory = ".\ExportedAddresses\"

If ($CreatePlaceHolder.IsPresent) {

    If (((Get-DistributionGroup $Group -ErrorAction 'SilentlyContinue').IsValid) -eq $true) {

        $OldDG = Get-DistributionGroup $Group

        [System.IO.Path]::GetInvalidFileNameChars() | ForEach {$Group = $Group.Replace($_,'_')}
        
        $OldName = [string]$OldDG.Name
        $OldDisplayName = [string]$OldDG.DisplayName
        $OldPrimarySmtpAddress = [string]$OldDG.PrimarySmtpAddress
        $OldAlias = [string]$OldDG.Alias
        $OldMembers = (Get-DistributionGroupMember $OldDG.Name).Name

        If(!(Test-Path -Path $ExportDirectory )){
            Write-Host "  Creating Directory: $ExportDirectory"
            New-Item -ItemType directory -Path $ExportDirectory | Out-Null
        }

        "EmailAddress" > "$ExportDirectory\$Group.csv"
        $OldDG.EmailAddresses >> "$ExportDirectory\$Group.csv"
        "x500:"+$OldDG.LegacyExchangeDN >> "$ExportDirectory\$Group.csv"

        Write-Host "  Creating Group: Cloud-$OldDisplayName"
    
        New-DistributionGroup `
            -Name "Cloud-$OldName" `
            -Alias "Cloud-$OldAlias" `
            -DisplayName "Cloud-$OldDisplayName" `
            -ManagedBy $OldDG.ManagedBy `
            -Members $OldMembers `
            -PrimarySmtpAddress "Cloud-$OldPrimarySmtpAddress" | Out-Null

        Sleep -Seconds 3

        Write-Host "  Setting Values For: Cloud-$OldDisplayName"

        Set-DistributionGroup `
            -Identity "Cloud-$OldName" `
            -AcceptMessagesOnlyFromSendersOrMembers $OldDG.AcceptMessagesOnlyFromSendersOrMembers `
            -RejectMessagesFromSendersOrMembers $OldDG.RejectMessagesFromSendersOrMembers `

        Set-DistributionGroup `
            -Identity "Cloud-$OldName" `
            -AcceptMessagesOnlyFrom $OldDG.AcceptMessagesOnlyFrom `
            -AcceptMessagesOnlyFromDLMembers $OldDG.AcceptMessagesOnlyFromDLMembers `
            -BypassModerationFromSendersOrMembers $OldDG.BypassModerationFromSendersOrMembers `
            -BypassNestedModerationEnabled $OldDG.BypassNestedModerationEnabled `
            -CustomAttribute1 $OldDG.CustomAttribute1 `
            -CustomAttribute2 $OldDG.CustomAttribute2 `
            -CustomAttribute3 $OldDG.CustomAttribute3 `
            -CustomAttribute4 $OldDG.CustomAttribute4 `
            -CustomAttribute5 $OldDG.CustomAttribute5 `
            -CustomAttribute6 $OldDG.CustomAttribute6 `
            -CustomAttribute7 $OldDG.CustomAttribute7 `
            -CustomAttribute8 $OldDG.CustomAttribute8 `
            -CustomAttribute9 $OldDG.CustomAttribute9 `
            -CustomAttribute10 $OldDG.CustomAttribute10 `
            -CustomAttribute11 $OldDG.CustomAttribute11 `
            -CustomAttribute12 $OldDG.CustomAttribute12 `
            -CustomAttribute13 $OldDG.CustomAttribute13 `
            -CustomAttribute14 $OldDG.CustomAttribute14 `
            -CustomAttribute15 $OldDG.CustomAttribute15 `
            -ExtensionCustomAttribute1 $OldDG.ExtensionCustomAttribute1 `
            -ExtensionCustomAttribute2 $OldDG.ExtensionCustomAttribute2 `
            -ExtensionCustomAttribute3 $OldDG.ExtensionCustomAttribute3 `
            -ExtensionCustomAttribute4 $OldDG.ExtensionCustomAttribute4 `
            -ExtensionCustomAttribute5 $OldDG.ExtensionCustomAttribute5 `
            -GrantSendOnBehalfTo $OldDG.GrantSendOnBehalfTo `
            -HiddenFromAddressListsEnabled $True `
            -MailTip $OldDG.MailTip `
            -MailTipTranslations $OldDG.MailTipTranslations `
            -MemberDepartRestriction $OldDG.MemberDepartRestriction `
            -MemberJoinRestriction $OldDG.MemberJoinRestriction `
            -ModeratedBy $OldDG.ModeratedBy `
            -ModerationEnabled $OldDG.ModerationEnabled `
            -RejectMessagesFrom $OldDG.RejectMessagesFrom `
            -RejectMessagesFromDLMembers $OldDG.RejectMessagesFromDLMembers `
            -ReportToManagerEnabled $OldDG.ReportToManagerEnabled `
            -ReportToOriginatorEnabled $OldDG.ReportToOriginatorEnabled `
            -RequireSenderAuthenticationEnabled $OldDG.RequireSenderAuthenticationEnabled `
            -SendModerationNotifications $OldDG.SendModerationNotifications `
            -SendOofMessageToOriginatorEnabled $OldDG.SendOofMessageToOriginatorEnabled `
            -BypassSecurityGroupManagerCheck
    }                
    Else {
        Write-Host "  ERROR: The distribution group '$Group' was not found" -ForegroundColor Red
        Write-Host
    }
}
ElseIf ($Finalize.IsPresent) {

        $TempDG = Get-DistributionGroup "Cloud-$Group"
        $TempPrimarySmtpAddress = $TempDG.PrimarySmtpAddress

        [System.IO.Path]::GetInvalidFileNameChars() | ForEach {$Group = $Group.Replace($_,'_')}

        $OldAddresses = @(Import-Csv "$ExportDirectory\$Group.csv")
    
        $NewAddresses = $OldAddresses | ForEach {$_.EmailAddress.Replace("X500","x500")}

        $NewDGName = $TempDG.Name.Replace("Cloud-","")
        $NewDGDisplayName = $TempDG.DisplayName.Replace("Cloud-","")
        $NewDGAlias = $TempDG.Alias.Replace("Cloud-","")
        $NewPrimarySmtpAddress = ($NewAddresses | Where {$_ -clike "SMTP:*"}).Replace("SMTP:","")

        Set-DistributionGroup `
            -Identity $TempDG.Name `
            -Name $NewDGName `
            -Alias $NewDGAlias `
            -DisplayName $NewDGDisplayName `
            -PrimarySmtpAddress $NewPrimarySmtpAddress `
            -HiddenFromAddressListsEnabled $False `
            -BypassSecurityGroupManagerCheck

        Set-DistributionGroup `
            -Identity $NewDGName `
            -EmailAddresses @{Add=$NewAddresses} `
            -BypassSecurityGroupManagerCheck

        Set-DistributionGroup `
            -Identity $NewDGName `
            -EmailAddresses @{Remove=$TempPrimarySmtpAddress} `
            -BypassSecurityGroupManagerCheck
    }
Else {
        Write-Host "  ERROR: No options selected, please use '-CreatePlaceHolder' or '-Finalize'" -ForegroundColor Red
        Write-Host
}