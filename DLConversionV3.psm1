
#############################################################################################
# DISCLAIMER:																				#
#																							#
# THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT					#
# PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY				#
# OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT		#
# LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR		#
# PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS		#
# AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR			#
# ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE	#
# FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS	#
# PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)	#
# ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,		#
# EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES						#
#############################################################################################


Function Start-DistributionListMigrationV3 
{
    
    <#
    .SYNOPSIS

    This is the trigger function that begins the process of allowing an administrator to migrate a distribution list from
    on premises to Office 365.

    .DESCRIPTION

    Trigger function.

    .PARAMETER GROUPSMTPADDRESS

    *Required*
    This is the SMTP address of the group based on the MAIL field in Active Directory.

    .PARAMETER LOGFOLDERPATH 

    *Required*
    Defines the location of the storage for log folders, exports, and trace files.

    .OUTPUTS

    Logs all activities and backs up all original data to the log folder directory.
    Moves the distribution group from on premieses source of authority to office 365 source of authority.

    .NOTES

    The following blog posts maintain documentation regarding this module.

    https://timmcmic.wordpress.com.  

    Refer to the first pinned blog post that is the table of contents.

    
    .EXAMPLE


    .EXAMPLE


    .EXAMPLE


    #>

    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$groupSMTPAddress,
        #Define other mandatory parameters
        [Parameter(Mandatory = $true)]
        [string]$logFolderPath,
        #Local Active Director Domain Controller Parameters
        [Parameter(Mandatory = $true)]
        [string]$globalCatalogServer,
        [Parameter(Mandatory = $true)]
        [pscredential]$activeDirectoryCredential,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Basic","Negotiate")]
        $activeDirectoryAuthenticationMethod="Negotiate",
        #Azure Active Directory Connect Parameters
        [Parameter(Mandatory = $false)]
        [string]$aadConnectServer=$NULL,
        #Exchange Online Parameters
        [Parameter(Mandatory = $false)]
        [pscredential]$exchangeOnlineCredential=$NULL,
        [Parameter(Mandatory = $false)]
        [string]$exchangeOnlineCertificateThumbPrint="",
        [Parameter(Mandatory = $false)]
        [string]$exchangeOnlineOrganizationName="",
        [Parameter(Mandatory = $false)]
        [ValidateSet("O365Default","O365GermanyCloud","O365China","O365USGovGCCHigh","O365USGovDoD")]
        [string]$exchangeOnlineEnvironmentName="O365Default",
        [Parameter(Mandatory = $false)]
        [string]$exchangeOnlineAppID="",
        #Define Microsoft Graph Parameters
        [Parameter(Mandatory = $false)]
        [ValidateSet("China","Global","USGov","USGovDod")]
        [string]$msGraphEnvironmentName="Global",
        [Parameter(Mandatory=$true)]
        [string]$msGraphTenantID="",
        [Parameter(Mandatory=$false)]
        [string]$msGraphCertificateThumbprint="",
        [Parameter(Mandatory=$false)]
        [string]$msGraphApplicationID="",
        [Parameter(Mandatory=$false)]
        [string]$msGraphClientSecret="",
        #Define other optional parameters
        [Parameter(Mandatory=$false)]
        [boolean]$overrideCentralizedMailTransportEnabled=$FALSE,
        [Parameter(Mandatory=$false)]
        [string]$customRoutingDomain="",
        [Parameter(Mandatory=$false)]
        [boolean]$testRecipientHealth=$true,
        [Parameter(Mandatory=$false)]
        [boolean]$testPropertyHealth=$true,
        #Define internal only paramters.
        [Parameter(Mandatory=$false)]
        [boolean]$isHealthCheck = $false,
        [Parameter(Mandatory =$FALSE)]
        [boolean]$allowTelemetryCollection=$TRUE,
        [Parameter(Mandatory =$FALSE)]
        [boolean]$allowDetailedTelemetryCollection=$TRUE,
        [Parameter(Mandatory =$FALSE)]
        [boolean]$iValidatedASupportedSyncVersion=$false
    )

    function generate-HTMLFile
    {
        #Prepare the HTML file for output.
        #Define the HTML file.

        out-logfile -string "Preparring to generate HTML file."

        $functionHTMLSuffix = "html"
        $global:functionHTMLFile = $global:LogFile.replace("log","$functionHTMLSuffix")

        out-logfile -string $global:functionHTMLFile
        $headerString = ("Migration Summary for "+$groupSMTPAddress)

        New-HTML -TitleText $groupSMTPAddress -FilePath $global:functionHTMLFile {
            New-HTMLHeader {
                New-HTMLText -Text $headerString -FontSize 24 -Color White -BackGroundColor Black -Alignment center
            }
            new-htmlMain{
                #Define HTML table options.

                New-HTMLTableOption -DataStore JavaScript

                if (($global:dlConversionV2Test.count -gt 0) -or ($global:preCreateErrors.count -gt 0) -or ($global:testOffice365Errors.count -gt 0) -or ($global:testOffice365PropertyErrors.count -gt 0))
                {
                    New-HTMLText -Text "Migration Errors Detected - Summary Information Below" -FontSize 24 -Color White -BackGroundColor RED -Alignment center

                    out-logfile -string "Generate Error Summary List"

                    New-HTMLSection -HeaderText "Error Count Summary" {
                        New-HTMLList{
                                new-htmlListItem -text ("Pre Office 365 Group Create Errors: "+$global:preCreateErrors.count) -fontSize 14
                                new-htmlListItem -text ("Test Office 365 Errors: "+$global:testOffice365Errors.count) -fontSize 14
                                new-htmlListItem -text ("Test Office 365 Property Errors: "+$global:testOffice365PropertyErrors.count) -fontSize 14
                                new-htmlListItem -text ("DLConversionV2 Dependency Count: "+$global:dlConversionV2Test.count) -fontSize 14
                            }
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red"  -CanCollapse -BorderRadius 10px -collapsed

                    out-logfile -string "Generate HTML for pre create errors."

                    if ($global:preCreateErrors.count -gt 0)
                    {
                        out-logfile -string "Precreate errors exist."

                        new-htmlSection -HeaderText ("Pre Office 365 Group Create Errors"){
                            new-htmlTable -DataTable ($global:preCreateErrors | select-object Alias,Name,PrimarySMTPAddressOrUPN,RecipientType,GroupType,RecipientOrUser,ExternalDirectoryObjectID,OnPremADAttribute,DN,isErrorMessage) -Filtering  {
                            } -AutoSize
                        } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red"  -CanCollapse -BorderRadius 10px -collapsed
                    }
                    else 
                    {
                        out-logfile -string "Precreate errors do not exist."
                    }

                    out-logfile -string "Generate HTML section for DLConversionV2 items."

                    if ($global:dlConversionV2Test.count -gt 0)
                    {
                        out-logfile -string "Precreate errors exist."

                        new-htmlSection -HeaderText ("DLConversionV2 Items on Group (Recommend Migration by DLConversionV2)"){
                            new-htmlTable -DataTable ($global:dlConversionV2Test | select-object Alias,Name,PrimarySMTPAddressOrUPN,RecipientType,GroupType,RecipientOrUser,ExternalDirectoryObjectID,OnPremADAttribute,DN,isErrorMessage) -Filtering  {
                            } -AutoSize
                        } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red"  -CanCollapse -BorderRadius 10px -collapsed
                    }
                    else 
                    {
                        out-logfile -string "Precreate errors do not exist."
                    }

                    out-logfile -string "Generate HTML for test office 365 errors."

                    if ($global:testOffice365Errors.count -gt 0)
                    {
                        out-logfile -string "Test Office 365 Errors exist."

                        new-htmlSection -HeaderText ("Test Office 365 Property Errors"){
                            new-htmlTable -DataTable ($global:testOffice365Errors | select-object Alias,Name,PrimarySMTPAddressOrUPN,RecipientType,GroupType,RecipientOrUser,ExternalDirectoryObjectID,OnPremADAttribute,DN,isErrorMessage) -Filtering  {
                            } -AutoSize
                        } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red"  -CanCollapse -BorderRadius 10px -collapsed
                    }
                    else 
                    {
                        out-logfile -string "Test Office 365 Errors do not exist."
                    }

                    out-logfile -string "Generate HTML for test office 365 errors."

                    if ($global:testOffice365PropertyErrors.count -gt 0)
                    {
                        out-logfile -string "Test Office 365 Errors exist."

                        new-htmlSection -HeaderText ("Test Office 365 Dependency Errors"){
                            new-htmlTable -DataTable ($global:testOffice365PropertyErrors | select-object Alias,Name,PrimarySMTPAddressOrUPN,RecipientType,GroupType,RecipientOrUser,ExternalDirectoryObjectID,OnPremADAttribute,DN,isErrorMessage) -Filtering  {
                            } -AutoSize
                        } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Red"  -CanCollapse -BorderRadius 10px -collapsed
                    }
                    else 
                    {
                        out-logfile -string "Test Office 365 Errors do not exist."
                    }
                }
                else 
                {
                    New-HTMLText -Text "*****MIGRATION SUCCESSFUL*****" -FontSize 24 -Color White -BackGroundColor Green -Alignment center
                }

                out-logfile -string "Generate HTML for Summary Counts."

                New-HTMLSection -HeaderText "Group Statistics" {
                    New-HTMLList{
                        new-htmlListItem -text ("The number of objects included in the member migration: "+$exchangeDLMembershipSMTP.count) -fontSize 14
                        new-htmlListItem -text ("The number of objects included in the reject memebers: "+$exchangeRejectMessagesSMTP.count) -fontSize 14
                        new-htmlListItem -text ("The number of objects included in the accept memebers: "+$exchangeAcceptMessagesSMTP.count) -fontSize 14
                        new-htmlListItem -text ("The number of objects included in the managedBY memebers: "+$exchangeManagedBySMTP.count) -fontSize 14
                        new-htmlListItem -text ("The number of objects included in the moderatedBY memebers: "+$exchangeModeratedBySMTP.count) -fontSize 14
                        new-htmlListItem -text ("The number of objects included in the bypassModeration memebers: "+$exchangeBypassModerationSMTP.count) -fontSize 14
                        new-htmlListItem -text ("The number of objects included in the grantSendOnBehalfTo memebers: "+$exchangeGrantSendOnBehalfToSMTP.count) -fontSize 14
                        new-htmlListItem -text ("The number of groups that the migrated DL is a member of = "+$allGroupsMemberOf.count) -fontSize 14
                        new-htmlListItem -text ("The number of groups that this group is a manager of: = "+$allGroupsManagedBy.count) -fontSize 14
                        new-htmlListItem -text ("The number of groups that this group has grant send on behalf to = "+$allGroupsGrantSendOnBehalfTo.count) -fontSize 14
                        new-htmlListItem -text ("The number of groups that have this group as bypass moderation = "+$allGroupsBypassModeration.count) -fontSize 14
                        new-htmlListItem -text ("The number of groups with accept permissions = "+$allGroupsAccept.count) -fontSize 14
                        new-htmlListItem -text ("The number of groups with reject permissions = "+$allGroupsReject.count) -fontSize 14
                        new-htmlListItem -text ("The number of mailboxes forwarding to this group is = "+$allUsersForwardingAddress.count) -fontSize 14
                        new-htmlListItem -text ("The number of groups this group is a co-manager on = "+$allGroupsCoManagedByBL.Count) -fontSize 14
                        new-htmlListItem -text ("The number of office 365 objects that the migrated DL is a member of = "+$allOffice365MemberOf.count) -fontSize 14
                        new-htmlListItem -text ("The number of office 365 objects that this group is a manager of: = "+$allOffice365ManagedBy.count) -fontSize 14
                        new-htmlListItem -text ("The number of office 365 objects that this group has grant send on behalf to = "+$allOffice365GrantSendOnBehalfTo.count) -fontSize 14
                        new-htmlListItem -text ("The number of office 365 objects that have this group as bypass moderation = "+$allOffice365BypassModeration.count) -fontSize 14
                        new-htmlListItem -text ("The number of office 365 objects with accept permissions = "+$allOffice365Accept.count) -fontSize 14
                        new-htmlListItem -text ("The number of office 365 objects with reject permissions = "+$allOffice365Reject.count) -fontSize 14
                        new-htmlListItem -text ("The number of office 365 mailboxes forwarding to this group is = "+$allOffice365ForwardingAddress.count) -fontSize 14
                        new-htmlListItem -text ("The number of recipients that have send as rights on the group to be migrated = "+$allOffice365SendAsAccessOnGroup.count) -fontSize 14
                        new-htmlListItem -text ("The number of office 365 recipients where the group has send as rights = "+$allOffice365SendAsAccess.count) -fontSize 14
                    }
                }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed

                out-logfile -string "Generate HTML for Telemetry Times"

                New-HTMLSection -HeaderText "Telemetry Time" {
                    New-HTMLList{
                        new-htmlListItem -text ("MigrationElapsedSeconds = "+$telemetryElapsedSeconds) -fontSize 14
                        new-htmlListItem -text ("TimeToNormalizeDNs = "+$telemetryNormalizeDN) -fontSize 14
                        new-htmlListItem -text ("TimeToValidateCloudRecipients = "+$telemetryValidateCloudRecipients) -fontSize 14
                        new-htmlListItem -text ("TimeToCollectOnPremDependency = "+$telemetryDependencyOnPrem) -fontSize 14
                        new-htmlListItem -text ("TimeToCollectOffice365Dependency = "+$telemetryCollectOffice365Dependency) -fontSize 14
                        new-htmlListItem -text ("TimeToConvertDLCloudOnly = "+$telemetryConvertGroupCloudOnly) -fontSize 14
                        new-htmlListItem -text ("TimeToConvertDLCloudOnlyExchangeOnline = "+$telemetryConvertGroupCloudOnlyExchangeOnline) -fontSize 14
                        new-htmlListItem -text ("TimeToCreateRoutingContact = "+$telemetryCreateRoutingContact) -fontSize 14
                    }
                }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed

                out-logfile -string "Generate HTML for Original DL Configuration"

                New-HTMLSection -HeaderText "Original DL Configuration (Active Directory)" {
                    New-HTMLList{
                        foreach ($object in $originalDLConfiguration.psObject.properties)
                        {
                            if ($object.Value.count -gt 1)
                            {
                                foreach ($value in $object.Value)
                                {
                                    $string = ($object.name + " " + $value.tostring())
                                    new-htmlListItem -text $string -fontSize 14
                                }
                            }
                            elseif ($object.value -ne $NULL)
                            {
                                $string = ($object.name + " " + $object.value.tostring())
                                new-htmlListItem -text $string -fontSize 14                            }
                            else
                            {
                                $string = ($object.name)
                                new-htmlListItem -text $string -fontSize 14
                            }
                        }
                    }
                }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed

                out-logfile -string "Generate HTML for Original Graph Configuration"

                New-HTMLSection -HeaderText "Original DL Configuration (Azure Active Directory)" {
                    New-HTMLList{
                        foreach ($object in $msGraphDLConfiguration.psObject.properties)
                        {
                            if ($object.Value.count -gt 1)
                            {
                                foreach ($value in $object.Value)
                                {
                                    $string = ($object.name + " " + $value.tostring())
                                    new-htmlListItem -text $string -fontSize 14
                                }
                            }
                            elseif ($object.value -ne $NULL)
                            {
                                $string = ($object.name + " " + $object.value.tostring())
                                new-htmlListItem -text $string -fontSize 14                            }
                            else
                            {
                                $string = ($object.name)
                                new-htmlListItem -text $string -fontSize 14
                            }
                        }
                    }
                }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed

                out-logfile -string "Generate HTML for Office 365 DL Configuration"

                New-HTMLSection -HeaderText "Original DL Configuration (Exchange Online)" {
                    New-HTMLList{
                        foreach ($object in $office365DLConfiguration.psObject.properties)
                        {
                            if ($object.Value.count -gt 1)
                            {
                                foreach ($value in $object.Value)
                                {
                                    $string = ($object.name + " " + $value.tostring())
                                    new-htmlListItem -text $string -fontSize 14
                                }
                            }
                            elseif ($object.value -ne $NULL)
                            {
                                $string = ($object.name + " " + $object.value.tostring())
                                new-htmlListItem -text $string -fontSize 14                            }
                            else
                            {
                                $string = ($object.name)
                                new-htmlListItem -text $string -fontSize 14
                            }
                        }
                    }
                }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed

                out-logfile -string "Generate HTML for Office 365 Group Configuration"

                New-HTMLSection -HeaderText "Original Group Configuration (Exchange Online)" {
                    New-HTMLList{
                        foreach ($object in  $office365GroupConfiguration.psObject.properties)
                        {
                            if ($object.Value.count -gt 1)
                            {
                                foreach ($value in $object.Value)
                                {
                                    $string = ($object.name + " " + $value.tostring())
                                    new-htmlListItem -text $string -fontSize 14
                                }
                            }
                            elseif ($object.value -ne $NULL)
                            {
                                $string = ($object.name + " " + $object.value.tostring())
                                new-htmlListItem -text $string -fontSize 14                            }
                            else
                            {
                                $string = ($object.name)
                                new-htmlListItem -text $string -fontSize 14
                            }
                        }
                    }
                }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed

                out-logfile -string "Generate HTML for Office 365 DL Configuration Post Migration"

                if ($office365DLConfigurationPostMigration -ne $NULL)
                {
                    New-HTMLSection -HeaderText "Office 365 DL Configuration Post Migration (Exchange Online)" {
                        New-HTMLList{
                            foreach ($object in  $office365DLConfigurationPostMigration.psObject.properties)
                            {
                                if ($object.Value.count -gt 1)
                                {
                                    foreach ($value in $object.Value)
                                    {
                                        $string = ($object.name + " " + $value.tostring())
                                        new-htmlListItem -text $string -fontSize 14
                                    }
                                }
                                elseif ($object.value -ne $NULL)
                                {
                                    $string = ($object.name + " " + $object.value.tostring())
                                    new-htmlListItem -text $string -fontSize 14                            }
                                else
                                {
                                    $string = ($object.name)
                                    new-htmlListItem -text $string -fontSize 14
                                }
                            }
                        }
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                out-logfile -string "Generate HTML for Graph Group post migration."

                if ($msGraphDLConfigurationPostMigration -ne $NULL)
                {
                    New-HTMLSection -HeaderText "Office 365 DL Configuration Post Migration (Exchange Online)" {
                        New-HTMLList{
                            foreach ($object in $msGraphDLConfigurationPostMigration.psObject.properties)
                            {
                                if ($object.Value.count -gt 1)
                                {
                                    foreach ($value in $object.Value)
                                    {
                                        $string = ($object.name + " " + $value.tostring())
                                        new-htmlListItem -text $string -fontSize 14
                                    }
                                }
                                elseif ($object.value -ne $NULL)
                                {
                                    $string = ($object.name + " " + $object.value.tostring())
                                    new-htmlListItem -text $string -fontSize 14                            }
                                else
                                {
                                    $string = ($object.name)
                                    new-htmlListItem -text $string -fontSize 14
                                }
                            }
                        }
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                out-logfile -string "Generate HTML for on premsies group membership."

                if ($originalDLConfiguration.member.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group Membership"){
                        new-htmlTable -DataTable ($originalDLConfiguration.member) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                out-logfile -string "Generate HTML for MS Graph Group membership."

                if ($msGraphDLMembership.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Graph Group Membership"){
                        new-htmlTable -DataTable ($msGraphDlMembership | select-object ID) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
    
                }

                out-logfile -string "Generate HTML for Office 365 DL Membership"

                if ($office365DLMembership.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 DL Membership"){
                        new-htmlTable -DataTable ($office365DLMembership) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                out-logfile -string "Generate HTML for Office 365 DL Membership Post Migration"

                if ($office365DLMembershipPostMigration.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 DL Membership Post Migration"){
                        new-htmlTable -DataTable ($office365DLMembershipPostMigration) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                out-logfile -string "Generate HTML for all on premsies normalized attributes."

                if ($exchangeDLMembershipSMTP.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises DL Membership Normalized"){
                        new-htmlTable -DataTable ($exchangeDLMembershipSMTP | select-object PrimarySMTPAddressOrUPN,Alias,ExternalDirectoryObjectID,DN,isAlreadyMigrated,RecipientOrUser,OnPremADAttributeCommonName,OnPremADAttribute) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($exchangeRejectMessagesSMTP.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Reject Normalized"){
                        new-htmlTable -DataTable ($exchangeRejectMessagesSMTP | select-object PrimarySMTPAddressOrUPN,Alias,ExternalDirectoryObjectID,DN,isAlreadyMigrated,RecipientOrUser,OnPremADAttributeCommonName,OnPremADAttribute) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($exchangeAcceptMessagesSMTP.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Accept Normalized"){
                        new-htmlTable -DataTable ($exchangeAcceptMessagesSMTP | select-object PrimarySMTPAddressOrUPN,Alias,ExternalDirectoryObjectID,DN,isAlreadyMigrated,RecipientOrUser,OnPremADAttributeCommonName,OnPremADAttribute) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($exchangeManagedBySMTP.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Accept Normalized"){
                        new-htmlTable -DataTable ($exchangeManagedBySMTP | select-object PrimarySMTPAddressOrUPN,Alias,ExternalDirectoryObjectID,DN,isAlreadyMigrated,RecipientOrUser,OnPremADAttributeCommonName,OnPremADAttribute) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }


                if ($exchangeModeratedBySMTP.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises ModeratedBy Normalized"){
                        new-htmlTable -DataTable ($exchangeModeratedBySMTP | select-object PrimarySMTPAddressOrUPN,Alias,ExternalDirectoryObjectID,DN,isAlreadyMigrated,RecipientOrUser,OnPremADAttributeCommonName,OnPremADAttribute) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($exchangeBypassModerationSMTP.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises BypassModeration Normalized"){
                        new-htmlTable -DataTable ($exchangeBypassModerationSMTP | select-object PrimarySMTPAddressOrUPN,Alias,ExternalDirectoryObjectID,DN,isAlreadyMigrated,RecipientOrUser,OnPremADAttributeCommonName,OnPremADAttribute) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($exchangeGrantSendOnBehalfToSMTP.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises GrantSendOnBehalfTo Normalized"){
                        new-htmlTable -DataTable ($exchangeGrantSendOnBehalfToSMTP | select-object PrimarySMTPAddressOrUPN,Alias,ExternalDirectoryObjectID,DN,isAlreadyMigrated,RecipientOrUser,OnPremADAttributeCommonName,OnPremADAttribute) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                out-logfile -string "Generate HTML for all on premises dependencies."

                if ($allGroupsMemberOf.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group Member Of"){
                        new-htmlTable -DataTable ($allGroupsMemberOf) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allGroupsReject.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group Reject"){
                        new-htmlTable -DataTable ($allGroupsReject) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allGroupsAccept.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group Accept"){
                        new-htmlTable -DataTable ($allGroupsAccept) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allGroupsBypassModeration.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group ModeratedBy"){
                        new-htmlTable -DataTable ($allGroupsBypassModeration) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allUsersForwardingAddress.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group Forwarding On Objects"){
                        new-htmlTable -DataTable ($allUsersForwardingAddress) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allGroupsGrantSendOnBehalfTo.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group SendOnBehalf Objects"){
                        new-htmlTable -DataTable ($allGroupsGrantSendOnBehalfTo) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allGroupsManagedBy.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group ManagedBy Objects"){
                        new-htmlTable -DataTable ($allGroupsManagedBy) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allObjectsFullMailboxAccess.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group FullMailboxAccess Objects"){
                        new-htmlTable -DataTable ($allObjectsFullMailboxAccess) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allGroupsCoManagedByBL.count -gt 0)
                {
                    new-htmlSection -HeaderText ("On Premises Group CoManagedBy Objects"){
                        new-htmlTable -DataTable ($allGroupsCoManagedByBL) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                out-logfile -string "Generate HTML for all Office 365 dependencies."

                if ($allOffice365MemberOf.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 Member of Other Groups Objects"){
                        new-htmlTable -DataTable ($allOffice365MemberOf) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allOffice365Accept.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 Accept Groups Objects"){
                        new-htmlTable -DataTable ($allOffice365Accept) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allOffice365Accept.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 Accept Groups Objects"){
                        new-htmlTable -DataTable ($allOffice365Accept) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allOffice365Reject.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 Reject Groups Objects"){
                        new-htmlTable -DataTable ($allOffice365Reject) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allOffice365BypassModeration.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 BypassModeration Groups Objects"){
                        new-htmlTable -DataTable ($allOffice365BypassModeration) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allOffice365ManagedBy.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 ManagedBy Groups Objects"){
                        new-htmlTable -DataTable ($allOffice365ManagedBy) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allOffice365GrantSendOnBehalfTo.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 GrantSendOnBehalfTo Groups Objects"){
                        new-htmlTable -DataTable ($allOffice365GrantSendOnBehalfTo) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allOffice365ForwardingAddress.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 ForwardingAddress Groups Objects"){
                        new-htmlTable -DataTable ($allOffice365ForwardingAddress) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allOffice365SendAsAccess.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 SendAs on Other Groups Objects"){
                        new-htmlTable -DataTable ($allOffice365SendAsAccess) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if ($allOffice365SendAsAccessOnGroup.count -gt 0)
                {
                    new-htmlSection -HeaderText ("Office 365 SendAs On Group"){
                        new-htmlTable -DataTable ($allOffice365SendAsAccessOnGroup) -Filtering {
                        } -AutoSize
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                out-logfile -string "Record routing contact configuration."

                if ($routingContactConfiguration -ne $NULL)
                {
                    New-HTMLSection -HeaderText "Hybrid Routing Contact" {
                        New-HTMLList{
                            foreach ($object in  $routingContactConfiguration.psObject.properties)
                            {
                                if ($object.Value.count -gt 1)
                                {
                                    foreach ($value in $object.Value)
                                    {
                                        $string = ($object.name + " " + $value.tostring())
                                        new-htmlListItem -text $string -fontSize 14
                                    }
                                }
                                elseif ($object.value -ne $NULL)
                                {
                                    $string = ($object.name + " " + $object.value.tostring())
                                    new-htmlListItem -text $string -fontSize 14                            }
                                else
                                {
                                    $string = ($object.name)
                                    new-htmlListItem -text $string -fontSize 14
                                }
                            }
                        }
                    }-HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }

                if (($global:dlConversionV2Test.count -gt 0) -or ($global:preCreateErrors.count -gt 0) -or ($global:testOffice365Errors.count -gt 0) -or ($global:testOffice365PropertyErrors.count -gt 0))
                {
                    out-logfile -string "Do not generate timeline - there was a failure."
                }
                else 
                {
                    out-logfile -string "Generate timeline."

                    new-htmlSection -HeaderText ("Migration Timeline Highlights"){
                        new-HTMLTimeLIne {
                            new-HTMLTimeLineItem -HeadingText "Migration Start Time" -Date $htmlStartTime
                            new-HTMLTimeLineItem -HeadingText "Start Parameter Validation" -Date $htmlStartValidationTime
                            new-HTMLTimeLineItem -HeadingText "Start Powershell Session Initialization" -Date $htmlStartPowershellSessions
                            new-HTMLTimeLineItem -HeadingText "Capture On-Premises DL Information" -Date $htmlCaptureOnPremisesDLInfo
                            new-HTMLTimeLineItem -HeadingText "Capture Office 365 DL Information" -Date $htmlCaptureOffice365DLConfiguration
                            new-HTMLTimeLineItem -HeadingText "Capture Graph DL Information" -Date $htmlCaptureGraphDLConfiguration
                            new-HTMLTimeLineItem -HeadingText "Capture Graph DL Membership" -Date $htmlCaptureGraphDLMembership
                            new-HTMLTimeLineItem -HeadingText "Capture Office 365 DL Membership" -Date $htmlCaptureOffice365DLMembership
                            new-HTMLTimeLineItem -HeadingText "Start Cloud Group Validation" -Date $htmlStartGroupValidation
                            new-HTMLTimeLineItem -HeadingText "Start Attribute Normalization" -Date $htmlStartAttributeNormalization
                            new-HTMLTimeLineItem -HeadingText "Start OnPremises -> Cloud Validation" -Date $htmlStartCloudValidationOnPremises
                            new-HTMLTimeLineItem -HeadingText "Start OnPremises Property -> Cloud Validation" -Date $htmlStartCloudValidationOffice365
                            new-HTMLTimeLineItem -HeadingText "Start Capture On-Premises Dependencies" -Date $htmlCaptureOnPremisesDependencies
                            new-HTMLTimeLineItem -HeadingText "Start Capture Office 365 Dependencies" -Date $htmlRecordOffice365Dependencies
                            new-HTMLTimeLineItem -HeadingText "Set EntraID Group Cloud Only" -Date $htmlSetGroupCloudOnly
                            new-HTMLTimeLineItem -HeadingText "Set Exchange Online Group Cloud Only" -Date $htmlTestExchangeOnlineCloudOnly
                            new-HTMLTimeLineItem -HeadingText "Capture Office 365 DL Info Post Migration" -Date $htmlCaptureOffice365InfoPostMigration
                            new-HTMLTimeLineItem -HeadingText "Create Routing Contact" -Date $htmlCreateRoutingContact
                            new-HTMLTimeLineItem -HeadingText "END" -Date $htmlEndTime
                        }
                    } -HeaderTextAlignment "Left" -HeaderTextSize "16" -HeaderTextColor "White" -HeaderBackGroundColor "Black"  -CanCollapse -BorderRadius 10px -collapsed
                }
            }
        } -online -ShowHTML
    }

    #Estbalish the HTML reporting start time.
    $htmlStartTime = get-date

    #Establish the graph scopes required for the module.
    $msGraphScopesRequired = @("User.Read.All", "Group.Read.All","Group-OnPremisesSyncBehavior.ReadWrite.All")

    #Initialize telemetry collection.

    $appInsightAPIKey = "63d673af-33f4-401c-931e-f0b64a218d89"
    $traceModuleName = "DLConversion"

    if ($allowTelemetryCollection -eq $TRUE)
    {
        start-telemetryConfiguration -allowTelemetryCollection $allowTelemetryCollection -appInsightAPIKey $appInsightAPIKey -traceModuleName $traceModuleName
    }

    #Create telemetry values.

    $telemetryDLConversionV3Version = $NULL
    $telemetryExchangeOnlineVersion = $NULL
    $telemetryAzureADVersion = $NULL
    $telemetryMSGraphAuthentication = $NULL
    $telemetryMSGraphUsers = $NULL
    $telemetryMSGraphGroups = $NULL
    $telemetryActiveDirectoryVersion = $NULL
    $telemetryOSVersion = (Get-CimInstance Win32_OperatingSystem).version
    $telemetryStartTime = get-universalDateTime
    $telemetryEndTime = $NULL
    [double]$telemetryElapsedSeconds = 0
    $telemetryEventName = "Start-DistributionListMigration-V3"
    $telemetryFunctionStartTime=$NULL
    $telemetryFunctionEndTime=$NULL
    [double]$telemetryNormalizeDN=0
    [double]$telemetryValidateCloudRecipients=0
    [double]$telemetryDependencyOnPrem=0
    [double]$telemetryCollectOffice365Dependency=0
    [double]$telemetryTimeToRemoveDL=0
    [double]$telemetryCreateOffice365DL=0
    [double]$telemetryCreateOffice365DLFirstPass=0
    [double]$telemetryReplaceOnPremDependency=0
    [double]$telemetryReplaceOffice365Dependency=0
    [boolean]$telemetryError=$FALSE

    $windowTitle = ("Start-DistributionListMigrationV3 "+$groupSMTPAddress)
    $host.ui.RawUI.WindowTitle = $windowTitle

    $global:blogURL = "https://timmcmic.wordpress.com"

    #Define variables utilized in the core function that are not defined by parameters.

    $coreVariables = @{ 
        useOnPremisesExchange = @{ "Value" = $FALSE ; "Description" = "Boolean determines if Exchange on premises should be utilized" }
        exchangeOnPremisesPowershellSessionName = @{ "Value" = "ExchangeOnPremises" ; "Description" = "Static exchange on premises powershell session name" }
        ADGlobalCatalogPowershellSessionName = @{ "Value" = "ADGlobalCatalog" ; "Description" = "Static AD Domain controller powershell session name" }
        exchangeOnlinePowershellModuleName = @{ "Value" = "ExchangeOnlineManagement" ; "Description" = "Static Exchange Online powershell module name" }
        activeDirectoryPowershellModuleName = @{ "Value" = "ActiveDirectory" ; "Description" = "Static active directory powershell module name" }
        msGraphAuthenticationPowershellModuleName = @{ "Value" = "Microsoft.Graph.Authentication" ; "Description" = "Static ms graph powershell name authentication" }
        dlConversionPowershellModule = @{ "Value" = "DLConversionV3" ; "Description" = "Static dlConversionv2 powershell module name" }
        globalCatalogPort = @{ "Value" = ":3268" ; "Description" = "Global catalog port definition" }
        globalCatalogWithPort = @{ "Value" = ($globalCatalogServer+($corevariables.globalCatalogPort.value)) ; "Description" = "Global catalog server with port" }
        msGraphUsersPowershellModuleName = @{ "Value" = "Microsoft.Graph.Users" ; "Description" = "Static ms graph powershell name users" }
        msGraphGroupsPowershellModuleName = @{ "Value" = "Microsoft.Graph.Groups" ; "Description" = "Static ms graph powershell name groups" }
    }

    #The variables below are utilized to define working parameter sets.
    #Some variables are assigned to single values - since these will be utilized with functions that query or set information.

    $onPremADAttributes = @{
        onPremAcceptMessagesFromDLMembers = @{"Value" = "dlMemSubmitPerms" ; "Description" = "LDAP Attribute for Accept Messages from DL Members"}
        onPremAcceptMessagesFromDLMembersCommon = @{"Value" = "AcceptMessagesFromMembers" ; "Description" = "LDAP Attribute for Accept Messages from DL Members"}
        onPremRejectMessagesFromDLMembers = @{"Value" = "dlMemRejectPerms" ; "Description" = "LDAP Attribute for Reject Messages from DL Members"}
        onPremRejectMessagesFromDLMembersCommon = @{"Value" = "RejectMessagesFromMembers" ; "Description" = "LDAP Attribute for Reject Messages from DL Members"}
        onPremBypassModerationFromDL = @{"Value" = "msExchBypassModerationFromDLMembersLink" ; "Description" = "LDAP Attribute for Bypass Moderation from DL Members"}
        onPremBypassModerationFromDLCommon = @{"Value" = "BypassModerationFromSendersOrMembers" ; "Description" = "LDAP Attribute for Bypass Moderation from DL Members"}
        onPremForwardingAddress = @{"Value" = "altRecipient" ; "Description" = "LDAP Attribute for ForwardingAddress"}
        onPremForwardingAddressCommon = @{"Value" = "ForwardingAddress" ; "Description" = "LDAP Attribute for ForwardingAddress"}
        onPremGrantSendOnBehalfTo = @{"Value" = "publicDelegates" ; "Description" = "LDAP Attribute for Grant Send on Behalf To"}
        onPremGrantSendOnBehalfToCommon = @{"Value" = "GrantSendOnBehalfTo" ; "Description" = "LDAP Attribute for Grant Send on Behalf To"}
        onPremRejectMessagesFromSenders = @{"Value" = "unauthorig" ; "Description" = "LDAP Attribute for Reject Messages from Sender"}
        onPremRejectMessagesFromSendersCommon = @{"Value" = "RejectMessagesFromSenders" ; "Description" = "LDAP Attribute for Reject Messages from Sender"}
        onPremAcceptMessagesFromSenders = @{"Value" = "authOrig" ; "Description" = "LDAp Attribute for Accept Messages From Sender"} 
        onPremAcceptMessagesFromSendersCommon = @{"Value" = "AcceptMessagesFromSenders" ; "Description" = "LDAp Attribute for Accept Messages From Sender"} 
        onPremManagedBy = @{"Value" = "managedBy" ; "Description" = "LDAP Attribute for Managed By"}
        onPremManagedByCommon = @{"Value" = "ManagedBy" ; "Description" = "LDAP Attribute for Managed By"}
        onPremCoManagedBy = @{"Value" = "msExchCoManagedByLink" ; "Description" = "LDAP Attributes for Co Managers (Muiltivalued ManagedBy)"}
        onPremCoManagedByCommon = @{"Value" = "ManagedBy" ; "Description" = "LDAP Attributes for Co Managers (Muiltivalued ManagedBy)"}
        onPremModeratedBy = @{"Value" = "msExchModeratedByLink" ; "Description" = "LDAP Attrbitute for Moderated By"}
        onPremModeratedByCommon = @{"Value" = "ModeratedBy" ; "Description" = "LDAP Attrbitute for Moderated By"}
        onPremBypassModerationFromSenders = @{"Value" = "msExchBypassModerationLink" ; "Description" = "LDAP Attribute for Bypass Moderation from Senders"}
        onPremBypassModerationFromSendersCommon = @{"Value" = "BypassModerationFromSendersorMembers" ; "Description" = "LDAP Attribute for Bypass Moderation from Senders"}
        onPremMembers = @{"Value" = "member" ; "Description" = "LDAP Attribute for Distribution Group Members" }
        onPremMembersCommon = @{"Value" = "Member" ; "Description" = "LDAP Attribute for Distribution Group Members" }
        onPremForwardingAddressBL = @{"Value" = "altRecipientBL" ; "Description" = "LDAP Backlink Attribute for Forwarding Address"}
        onPremRejectMessagesFromDLMembersBL = @{"Value" = "dlMemRejectPermsBL" ; "Description" = "LDAP Backlink Attribute for Reject Messages from DL Members"}
        onPremAcceptMessagesFromDLMembersBL = @{"Value" = "dlMemSubmitPermsBL" ; "Description" = "LDAP Backlink Attribute for Accept Messages from DL Members"}
        onPremManagedObjects = @{"Value" = "managedObjects" ; "Description" = "LDAP Backlink Attribute for Managed By"}
        onPremMemberOf = @{"Value" = "memberOf" ; "Description" = "LDAP Backlink Attribute for Members"}
        onPremBypassModerationFromDLMembersBL = @{"Value" = "msExchBypassModerationFromDLMembersBL" ; "Description" = "LDAP Backlink Attribute for Bypass Moderation from DL Members"}
        onPremCoManagedByBL = @{"Value" = "msExchCoManagedObjectsBL" ; "Description" = "LDAP Backlink Attribute for Co Managers (Multivalued ManagedBY)"}
        onPremGrantSendOnBehalfToBL = @{"Value" = "publicDelegatesBL" ; "Description" = "LDAP Backlink Attribute for Grant Send On Behalf To"}
        onPremGroupType = @{"Value" = "groupType" ; "Description" = "Value representing universal / global / local / security / distribution"}
    }

    #Define the Office 365 attributes that will be used for filters.

    $office365Attributes  = @{ 
        office365AcceptMessagesFrom = @{ "Value" = "AcceptMessagesOnlyFromDLMembers" ; "Description" = "All Office 365 objects that have accept messages from senders or members for the migrated group"}
        office365BypassModerationFrom = @{ "Value" = "BypassModerationFromDLMembers" ; "Description" = "All Office 365 objects that have bypass moderation from senders or members for the migrated group"}
        office365CoManagers = @{ "Value" = "CoManagedBy" ; "Description" = "ALl office 365 objects that have managed by set for the migrated group"}
        office365GrantSendOnBehalfTo = @{ "Value" = "GrantSendOnBehalfTo" ; "Description" = "All Office 365 objects that have grant sent on behalf to for the migrated group"}
        office365ManagedBy = @{ "Value" = "ManagedBy" ; "Description" = "All Office 365 objects that have managed by set on the group"}
        office365Members = @{ "Value" = "Members" ; "Description" = "All Office 365 groups that the migrated group is a member of"}
        office365RejectMessagesFrom = @{ "Value" = "RejectMessagesFromDLMembers" ; "Description" = "All Office 365 groups that have the reject messages from senders or members right assignged to the migrated group"}
        office365ForwardingAddress = @{ "Value" = "ForwardingAddress" ; "Description" = "All Office 365 objects that have the migrated group set for forwarding"}
        office365BypassModerationusers = @{ "Value" = "BypassModerationFromSendersOrMembers" ; "Description" = "All Office 365 objects that have bypass moderation for the migrated group"}
        office365UnifiedAccept = @{ "Value" = "AcceptMessagesOnlyFromSendersOrMembers" ; "Description" = "All Office 365 Unified Groups that the migrated group has accept messages from senders or members rights assigned"}
        office365UnifiedReject = @{ "Value" = "RejectMessagesFromSendersOrMembers" ; "Description" = "All Office 365 Unified Groups that the migrated group has reject messages from senders or members rights assigned"}
    }

    #Define XML files to contain backups.

    $xmlFiles = @{
        originalDLConfigurationADXML = @{ "Value" =  "originalDLConfigurationADXML" ; "Description" = "XML file that exports the original DL configuration"}
        originalDLConfigurationUpdatedXML = @{ "Value" =  "originalDLConfigurationUpdatedXML" ; "Description" = "XML file that exports the updated DL configuration"}
        office365DLConfigurationXML = @{ "Value" =  "office365DLConfigurationXML" ; "Description" = "XML file that exports the Office 365 DL configuration"}
        office365GroupConfigurationXML = @{ "Value" = "office365GroupConfigurationXML" ; "Description" = "XML file that exports the Office 365 Group configuraiton"}
        office365DLConfigurationPostMigrationXML = @{ "Value" =  "office365DLConfigurationPostMigrationXML" ; "Description" = "XML file that exports the Office 365 DL configuration post migration"}
        office365DLMembershipPostMigrationXML = @{ "Value" =  "office365DLMembershipPostMigrationXML" ; "Description" = "XML file that exports the Office 365 DL membership post migration"}
        exchangeDLMembershipSMTPXML = @{ "Value" =  "exchangeDLMemberShipSMTPXML" ; "Description" = "XML file that holds the SMTP addresses of the on premises DL membership"}
        exchangeRejectMessagesSMTPXML = @{ "Value" =  "exchangeRejectMessagesSMTPXML" ; "Description" = "XML file that holds the Reject Messages From Senders or Members property of the on premises DL"}
        exchangeAcceptMessagesSMTPXML = @{ "Value" =  "exchangeAcceptMessagesSMTPXML" ; "Description" = "XML file that holds the Accept Messages from Senders or Members property of the on premises DL"}
        exchangeManagedBySMTPXML = @{ "Value" =  "exchangeManagedBySMTPXML" ; "Description" = "XML file that holds the ManagedBy proprty of the on premises DL"}
        exchangeModeratedBySMTPXML = @{ "Value" =  "exchangeModeratedBYSMTPXML" ; "Description" = "XML file that holds the Moderated By property of the on premises DL"}
        exchangeBypassModerationSMTPXML = @{ "Value" =  "exchangeBypassModerationSMTPXML" ; "Description" = "XML file that holds the Bypass Moderation From Senders or Members property of the on premises DL"}
        exchangeGrantSendOnBehalfToSMTPXML = @{ "Value" =  "exchangeGrantSendOnBehalfToXML" ; "Description" = "XML file that holds the Grant Send On Behalf To property of the on premises DL"}
        exchangeSendAsSMTPXML = @{ "Value" =  "exchangeSendASSMTPXML" ; "Description" = "XML file that holds the Send As rights of the on premises DL"}
        allGroupsMemberOfXML = @{ "Value" =  "allGroupsMemberOfXML" ; "Description" = "XML file that holds all of on premises groups the migrated group is a member of"}
        allGroupsRejectXML = @{ "Value" =  "allGroupsRejectXML" ; "Description" = "XML file that holds all of the on premises groups the migrated group has reject rights assigned"}
        allGroupsAcceptXML = @{ "Value" =  "allGroupsAcceptXML" ; "Description" = "XML file that holds all of the on premises groups the migrated group has accept rights assigned"}
        allGroupsBypassModerationXML = @{ "Value" =  "allGroupsBypassModerationXML" ; "Description" = "XML file that holds all of the on premises groups that the migrated group has bypass moderation rights assigned"}
        allUsersForwardingAddressXML = @{ "Value" =  "allUsersForwardingAddressXML" ; "Description" = "XML file that holds all recipients the migrated group hsa forwarding address set on"}
        allGroupsGrantSendOnBehalfToXML = @{ "Value" =  "allGroupsGrantSendOnBehalfToXML" ; "Description" = "XML file that holds all of the on premises objects that the migrated group hsa grant send on behalf to on"}
        allGroupsManagedByXML = @{ "Value" =  "allGroupsManagedByXML" ; "Description" = "XML file that holds all of the on premises objects the migrated group has managed by rights assigned"}
        allGroupsSendAsXML = @{ "Value" =  "allGroupSendAsXML" ; "Description" = "XML file that holds all of the on premises objects that have the migrated group with send as rights assigned"}
        allGroupsSendAsNormalizedXML= @{ "Value" = "allGroupsSendAsNormalizedXML" ; "Description" = "XML file that holds all normalized send as right"}
        allGroupsFullMailboxAccessXML = @{ "Value" =  "allGroupsFullMailboxAccessXML" ; "Description" = "XML file that holds all full mailbox access rights assigned to the migrated group"}
        allMailboxesFolderPermissionsXML = @{ "Value" =  "allMailboxesFolderPermissionsXML" ; "Description" = "XML file that holds all mailbox folder permissions assigned to the migrated group"}
        allOffice365MemberOfXML= @{ "Value" = "allOffice365MemberOfXML" ; "Description" = "XML file that holds All cloud only groups that have the migrated group as a member"}
        allOffice365AcceptXML= @{ "Value" = "allOffice365AcceptXML" ; "Description" = "XML file that holds All cloud only groups that have the migrated group assigned accept messages from senders or members rights"}
        allOffice365RejectXML= @{ "Value" = "allOffice365RejectXML" ; "Description" = "XML file that holds All cloud only groups that have the migrated group assigned reject messages from senders or members rights"}
        allOffice365BypassModerationXML= @{ "Value" = "allOffice365BypassModerationXML" ; "Description" = "XML file that holds All cloud only groups that have the migrated group assigned bypass moderation from senders or members"}
        allOffice365GrantSendOnBehalfToXML= @{ "Value" = "allOffice365GrantSentOnBehalfToXML" ; "Description" = "XML file that holds All cloud only groups that have the migrated group assigned grant send on behalf to rights"}
        allOffice365ManagedByXML= @{ "Value" = "allOffice365ManagedByXML" ; "Description" = "XML file that holds All cloud only groups that have the migrated group assigned managed by rights"}
        allOffice365ForwardingAddressXML= @{ "Value" = "allOffice365ForwardingAddressXML" ; "Description" = " XML file that holds all cloud only recipients where forwarding is set to the migrated grouop"}
        allOffic365SendAsAccessXML = @{ "Value" =  "allOffice365SendAsAccessXML" ; "Description" = "XML file that holds all cloud groups where send as rights are assigned to the migrated group"}
        allOffice365FullMailboxAccessXML = @{ "Value" =  "allOffice365FullMailboxAccessXML" ; "Description" = "XML file that holds all cloud only objects where full mailbox access is assigned to the migrated group"}
        allOffice365MailboxesFolderPermissionsXML = @{ "Value" =  'allOffice365MailboxesFolderPermissionsXML' ; "Description" = "XML file that holds all cloud only recipients where a mailbox folder permission is assigned to the migrated group"}
        allOffice365SendAsAccessOnGroupXML = @{ "Value" =  'allOffice365SendAsAccessOnGroupXML' ; "Description" = "XML file that holds all cloud only send as rights assigned to the migrated group"}
        routingContactXML= @{ "Value" = "routingContactXML" ; "Description" = "XML file holds the routing contact configuration when intially created"}
        routingContactUpdatedXML= @{ "Value" = "routingContactUpdatedXML" ; "Description" = "XML file holds the routing contact configuration when intially created"}
        routingDynamicGroupXML= @{ "Value" = "routingDynamicGroupXML" ; "Description" = "XML file holds the routing contact configuration when mail enabled"}
        allGroupsCoManagedByXML= @{ "Value" = "allGroupsCoManagedByXML" ; "Description" = "XML file holds all on premises objects that the migrated group has managed by rights assigned"}
        retainOffice365RecipientFullMailboxAccessXML= @{ "Value" = "office365RecipientFullMailboxAccess.xml" ; "Description" = "Import XML file for pre-gathered full mailbox access rights in Office 365"}
        retainMailboxFolderPermsOffice365XML= @{ "Value" = "office365MailboxFolderPermissions.xml" ; "Description" = "Import XML file for pre-gathered mailbox folder permissions in Office 365"}
        retainOnPremRecipientFullMailboxAccessXML= @{ "Value" = "onPremRecipientFullMailboxAccess.xml" ; "Description" = "Import XML for pre-gathered full mailbox access rights "}
        retainOnPremMailboxFolderPermissionsXML= @{ "Value" = "onPremailboxFolderPermissions.xml" ; "Description" = "Import XML file for mailbox folder permissions"}
        retainOnPremRecipientSendAsXML= @{ "Value" = "onPremRecipientSendAs.xml" ; "Description" = "Import XML file for send as permissions"}
        azureDLConfigurationXML = @{"Value" = "azureADDL" ; "Description" = "Export XML file holding the configuration from azure active directory"}
        azureDLMembershipXML = @{"Value" = "azureADDLMembership" ; "Description" = "Export XML file holding the membership of the Azure AD group"}
        msGraphDLConfigurationXML = @{"Value" = "msGraphADDL" ; "Description" = "Export XML file holding the configuration from azure active directory"}
        msGraphDLConfigurationPostMigrationXML = @{"Value" = "msGraphADDLPostMigration" ; "Description" = "Export XML file holding the configuration from azure active directory"}
        msGraphDLMembershipXML = @{"Value" = "msGraphADDLMembership" ; "Description" = "Export XML file holding the membership of the Azure AD group"}
        preCreateErrorsXML = @{"value" = "preCreateErrors" ; "Description" = "Export XML of all precreate errors for group to be migrated."}
        testOffice365ErrorsXML = @{"value" = "testOffice365Errors" ; "Description" = "Export XML of all tested recipient errors in Offic3 365."}
        testOffice365PropertyErrorsXML = @{"value" = "testOffice365PropertyErrors" ; "Description" = "Export XML of all property errors in Office 365."}
        office365DLMembership = @{"Value" = "office365DLMembership" ; "Description" = "Original Office 365 DL Membership"}
    }

    #On premises variables for the distribution list to be migrated.

    $originalDLConfiguration=$NULL #This holds the on premises DL configuration for the group to be migrated.
    $originalAzureADConfiguration=$NULL #This holds the azure ad DL configuration
    $originalDLConfigurationUpdated=$NULL #This holds the on premises DL configuration post the rename operations.
    $routingContactConfig=$NULL #Holds the mail routing contact configuration.
    $routingDynamicGroupConfig=$NULL #Holds the dynamic distribution list configuration used for mail routing.
    $routingContactConfiguration=$NULL #This is the empty routing contact configuration.
    #[array]$exchangeDLMembershipSMTP=@() #Array of DL membership from AD.
    $exchangeDLMembershipSMTP = New-Object Syste m.Collections.ArrayList
    #[array]$exchangeRejectMessagesSMTP=@() #Array of members with reject permissions from AD.
    $exchangeRejectMessagesSMTP = New-Object System.Collections.ArrayList
    #[array]$exchangeAcceptMessagesSMTP=@() #Array of members with accept permissions from AD.
    $exchangeAcceptMessagesSMTP = New-Object System.Collections.ArrayList
    #[array]$exchangeManagedBySMTP=@() #Array of members with manage by rights from AD.
    $exchangeManagedBySMTP = New-Object System.Collections.ArrayList
    #[array]$exchangeModeratedBySMTP=@() #Array of members  with moderation rights.
    $exchangeModeratedBySMTP = New-Object System.Collections.ArrayList
    #[array]$exchangeBypassModerationSMTP=@() #Array of objects with bypass moderation rights from AD.
    $exchangeBypassModerationSMTP = New-Object System.Collections.ArrayList
    #[array]$exchangeGrantSendOnBehalfToSMTP=@() #Array of objects with grant send on behalf to normalized SMTP
    $exchangeGrantSendOnBehalfToSMTP = New-Object System.Collections.ArrayList
    #[array]$exchangeSendAsSMTP=@() #Array of objects wtih send as rights normalized SMTP
    $exchangeSendAsSMTP = New-Object System.Collections.ArrayList
    #The following variables hold information regarding other groups in the environment that have dependnecies on the group to be migrated.

    [array]$allGroupsMemberOf=$NULL #Complete AD information for all groups the migrated group is a member of.
    [array]$allGroupsReject=$NULL #Complete AD inforomation for all groups that the migrated group has reject mesages from.
    [array]$allGroupsAccept=$NULL #Complete AD information for all groups that the migrated group has accept messages from.
    [array]$allGroupsBypassModeration=$NULL #Complete AD information for all groups that the migrated group has bypass moderations.
    [array]$allUsersForwardingAddress=$NULL #All users on premsies that have this group as a forwarding DN.
    [array]$allGroupsGrantSendOnBehalfTo=$NULL #All dependencies on premsies that have grant send on behalf to.
    [array]$allGroupsManagedBy=$NULL #All dependencies on premises that have managed by rights
    [array]$allObjectsFullMailboxAccess=$NULL #All dependencies on premises that have full mailbox access rights
    [array]$allObjectSendAsAccess=$NULL #All dependencies on premises that have the migrated group with send as rights.
    [array]$allObjectsSendAsAccessNormalized=@() #All dependencies send as rights normalized
    [array]$allMailboxesFolderPermissions=@() #All dependencies on premises with mailbox folder permissions defined
    [array]$allGroupsCoManagedByBL=$NULL #All groups on premises where the migrated group is a manager

    #The following variables hold information regarding Office 365 objects that have dependencies on the migrated DL.

    [array]$allOffice365MemberOf=$NULL #All cloud only groups the migrated group is a member of.
    [array]$allOffice365Accept=$NULL #All cloud only groups the migrated group has accept messages from senders or members.
    [array]$allOffice365Reject=$NULL #All cloud only groups the migrated group has reject messages from senders or members.
    [array]$allOffice365BypassModeration=$NULL #All cloud only groups the migrated group has bypass moderation from senders or members.
    [array]$allOffice365ManagedBy=$NULL #All cloud only groups the migrated group has managed by rights on.
    [array]$allOffice365GrantSendOnBehalfTo=$NULL #All cloud only groups the migrated group has grant send on behalf to on.
    [array]$allOffice365ForwardingAddress=$NULL #All cloud only recipients the migrated group has forwarding address 
    [array]$allOffice365FullMailboxAccess=$NULL #All cloud only recipients the migrated group has full ,amilbox access on.
    [array]$allOffice365SendAsAccess=$NULL #All cloud only groups the migrated group has send as access on.
    [array]$allOffice365SendAsAccessOnGroup = $NULL #All send as permissions set on the on premises group that are set in the cloud.
    [array]$allOffice365MailboxFolderPermissions=$NULL #All cloud only groups the migrated group has mailbox folder permissions on.

    #Cloud variables for the distribution list to be migrated.

    $office365DLConfiguration = $NULL #This holds the office 365 DL configuration for the group to be migrated.
    $office365GroupConfiguration = $NULL #This holds the office 365 group configuration for the group to be migrated.
    $msGraphDLConfiguration = $NULL #This holds the Azure AD DL configuration
    $msGraphDlMembership = $NULL
    $office365DLConfigurationPostMigration = $NULL #This hold the Office 365 DL configuration post migration.
    $msGraphDLConfigurationPostMigration = $NULL
    $office365DLMembership=$NULL
    $office365DLMembershipPostMigration=$NULL #This holds the Office 365 DL membership information post migration
    $msGraphURL = ""

    $dlPropertySet = '*' #Clear all properties of a given object

    [array]$global:preCreateErrors=@()
    [array]$global:testOffice365Errors=@()
    [array]$global:testOffice365PropertyErrors=@()
    [array]$global:generalErrors=@()
    [string]$isTestError="No"

    #Initilize the log file.

    $global:logFile=$NULL #This is the global variable for the calculated log file name
    [string]$global:staticFolderName="\DLMigration\"

    if ($isHealthCheck -eq $false)
    {
        new-LogFile -groupSMTPAddress $groupSMTPAddress.trim() -logFolderPath $logFolderPath
        $traceFilePath = $logFolderPath + $global:staticFolderName
        out-logfile -string ("Log File: "+$global:logFile)
        out-logfile -string ("Trace File: "+$traceFilePath)
    }

    if ($iValidatedASupportedSyncVersion -eq $FALSE)
    {
        out-logfile -string "It is a requirement to set this value to TRUE to migrate."
        out-logfile -string "Setting this value to true signifies the following:"
        out-logfile -string "All Entra Connect Servers are running version 2.5.76.0 or newer."
        out-logfile -string "All Entra Cloud Sync Agents are running version 1.1.1373.0 or newer."
        out-logfile -string "Failure to utilize a support version may result in migrations being reverted and lose of changes in EntraID / M365"
        out-logfile -string "EXCEPTION_DID_NOT_VALIDATE_SUPPORTED_SYNC_VERSIONS" -isError:$true
    }

    $htmlFunctionStartTime = get-Date

    out-logfile -string "********************************************************************************"
    out-logfile -string "NOTICE"
    out-logfile -string "Telemetry collection is now enabled by default."
    out-logfile -string "For information regarding telemetry collection see https://timmcmic.wordpress.com/2022/11/14/4288/"
    out-logfile -string "Administrators may opt out of telemetry collection by using -allowTelemetryCollection value FALSE"
    out-logfile -string "Telemetry collection is appreciated as it allows further development and script enhancement."
    out-logfile -string "********************************************************************************"

    #Output all parameters bound or unbound and their associated values.

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "PARAMETERS"
    Out-LogFile -string "********************************************************************************"

    write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

    Out-LogFile -string "================================================================================"
    Out-LogFile -string "BEGIN START-DISTRIBUTIONLISTMIGRATIONV3"
    Out-LogFile -string "================================================================================"

    out-logfile -string ("Runtime start UTC: " + $telemetryStartTime.ToString())

    if ($errorActionPreference -ne "Continue")
    {
        out-logfile -string ("Current Error Action Preference: "+$errorActionPreference)
        $errorActionPreference = "Continue"
        out-logfile -string ("New Error Action Preference: "+$errorActionPreference)
    }
    else
    {
        out-logfile -string ("Current Error Action Preference is CONTINUE: "+$errorActionPreference)
    }

    out-logfile -string "Obtain the graph URL for calls."

    $msGraphURL = get-GraphEnvironment -msGraphEnvironmentName $msGraphEnvironmentName -useBeta:$true

    out-logfile -string ("Graph URL: "+$msGraphURL)

    $htmlStartValidationTime = get-date

    $groupSMTPAddress = remove-stringSpace -stringToFix $groupSMTPAddress
    $globalCatalogServer = remove-stringSpace -stringToFix $globalCatalogServer
    $logFolderPath = remove-stringSpace -stringToFix $logFolderPath 

    if ($aadConnectServer -ne $NULL)
    {
        $aadConnectServer = remove-stringSpace -stringToFix $aadConnectServer
    }

    if ($exchangeOnlineCertificateThumbPrint -ne "")
    {
        $exchangeOnlineCertificateThumbPrint=remove-stringSpace -stringToFix $exchangeOnlineCertificateThumbPrint
    }

    $exchangeOnlineEnvironmentName=remove-stringSpace -stringToFix $exchangeOnlineEnvironmentName

    if ($exchangeOnlineOrganizationName -ne "")
    {
        $exchangeOnlineOrganizationName=remove-stringSpace -stringToFix $exchangeOnlineOrganizationName
    }

    if ($exchangeOnlineAppID -ne "")
    {
        $exchangeOnlineAppID=remove-stringSpace -stringToFix $exchangeOnlineAppID
    }

    $exchangeAuthenticationMethod=remove-StringSpace -stringToFix $exchangeAuthenticationMethod

    $msGraphTenantID = remove-stringSpace -stringToFix $msGraphTenantID
    $msGraphCertificateThumbprint = remove-stringSpace -stringToFix $msGraphCertificateThumbprint
    $msGraphApplicationID = remove-stringSpace -stringToFix $msGraphApplicationID
    $msGraphClientSecret = remove-stringSpace -stringToFix $msGraphClientSecret

    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string " RECORD VARIABLES"
    Out-LogFile -string "********************************************************************************"

    foreach ($dlProperty in $dlPropertySet)
    {
        Out-LogFile -string $dlProperty
    }

    Out-LogFile -string ("DL property set to be cleared legacy = ")

    foreach ($dlProperty in $dlPropertiesToClearLegacy)
    {
        Out-LogFile -string $dlProperty
    }

    Out-LogFile -string ("DL property set to be cleared modern = ")

    foreach ($dlProperty in $dlPropertiesToClearModern)
    {
        Out-LogFile -string $dlProperty
    }

    out-logfile -string ("Exchange legacy schema version: "+$exchangeLegacySchemaVersion)

    write-hashTable -hashTable $xmlFiles
    write-hashTable -hashTable $office365Attributes
    write-hashTable -hashTable $onPremADAttributes
    write-hashTable -hashTable $coreVariables

    Out-LogFile -string "********************************************************************************"

    #Perform paramter validation manually.

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "ENTERING PARAMTER VALIDATION"
    Out-LogFile -string "********************************************************************************"

    #Validate any credentials passed are of type PS Credential

    out-logfile -string "Testing global catalog credentials"
    test-credentials -credentialsToTest $activeDirectoryCredential

    #Validate Exchange Online Credentials

    Out-LogFile -string "Validating Exchange Online Credentials."

    start-parameterValidationExchange -exchangeOnlineCredential $exchangeOnlineCredential -exchangeOnlineCertificateThumbprint $exchangeOnlineCertificateThumbprint -exchangeOnlineOrganizationName $exchangeOnlineOrganizationName -exchangeOnlineAppID $exchangeOnlineAppID

    Out-LogFile -string "Validating Exchange Online Credentials."

    start-parameterValidationGraph -msGraphCertificateThumbPrint $msGraphCertificateThumbprint -msGraphTenantID $msGraphTenantID -msGraphApplicationID $msGraphApplicationID -msGraphClientSecret $msGraphClientSecret

    Out-LogFile -string "END PARAMETER VALIDATION"
    Out-LogFile -string "********************************************************************************"

    $htmlStartPowershellSessions = get-date

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "ESTABLISH POWERSHELL SESSIONS"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "Calling Test-PowerShellModule to validate the Exchange Module is installed."

    $telemetryExchangeOnlineVersion = Test-PowershellModule -powershellModuleName $corevariables.exchangeOnlinePowershellModuleName.value -powershellVersionTest:$TRUE

    Out-LogFile -string "Calling Test-PowerShellModule to validate the Active Directory is installed."

    $telemetryActiveDirectoryVersion = Test-PowershellModule -powershellModuleName $corevariables.activeDirectoryPowershellModuleName.value

    out-logfile -string "Calling Test-PowershellModule to validate the DL Conversion Module version installed."

    $telemetryDLConversionV3Version = Test-PowershellModule -powershellModuleName $corevariables.dlConversionPowershellModule.value -powershellVersionTest:$TRUE

    out-logfile -string "Calling Test-PowershellModule to validate the Microsoft Graph Authentication versions installed."

    $telemetryMSGraphAuthentication = test-powershellModule -powershellmodulename $corevariables.msgraphauthenticationpowershellmodulename.value -powershellVersionTest:$TRUE

    out-logfile -string "Calling Test-PowershellModule to validate the Microsoft Graph Users versions installed."

    $telemetryMSGraphUsers = test-powershellModule -powershellmodulename $corevariables.msgraphuserspowershellmodulename.value -powershellVersionTest:$TRUE

   out-logfile -string "Calling Test-PowershellModule to validate the Microsoft Graph Users versions installed."

    $telemetryMSGraphGroups = test-powershellModule -powershellmodulename $corevariables.msgraphgroupspowershellmodulename.value -powershellVersionTest:$TRUE

    Out-LogFile -string "Calling New-ExchangeOnlinePowershellSession to create session to office 365."

    New-ExchangeOnlinePowershellSession -exchangeOnlineCredentials $exchangeOnlineCredential -exchangeOnlineEnvironmentName $exchangeOnlineEnvironmentName -exchangeOnlineAppID $exchangeOnlineAppID -exchangeOnlineOrganizationName $exchangeOnlineOrganizationName -exchangeOnlineCertificateThumbPrint $exchangeOnlineCertificateThumbPrint -debugLogPath $traceFilePath

    Out-LogFile -string "Calling new-msGraphPowershellSession to create new connection to msGraph active directory."

    new-msGraphPowershellSession -msGraphCertificateThumbprint $msGraphCertificateThumbprint -msGraphApplicationID $msGraphApplicationID -msGraphTenantID $msGraphTenantID -msGraphEnvironmentName $msGraphEnvironmentName -msGraphScopesRequired $msGraphScopesRequired -msGraphClientSecret $msGraphClientSecret

    validate-GraphContext -msGraphScopesRequired $msGraphScopesRequired

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END ESTABLISH POWERSHELL SESSIONS"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN GET ORIGINAL DL CONFIGURATION LOCAL AND CLOUD"
    Out-LogFile -string "********************************************************************************"

    $htmlCaptureOnPremisesDLInfo = get-date

    #At this point we are ready to capture the original DL configuration.  We'll use the ad provider to gather this information.

    $originalDLConfiguration = Get-ADObjectConfiguration -groupSMTPAddress $groupSMTPAddress -globalCatalogServer $corevariables.globalCatalogWithPort.value -parameterSet $dlPropertySet -errorAction STOP -adCredential $activeDirectoryCredential -isGroupTest $TRUE

    Out-LogFile -string "Create an XML file backup of the on premises DL Configuration"

    Out-XMLFile -itemToExport $originalDLConfiguration -itemNameToExport $xmlFiles.originalDLConfigurationADXML.value

    $htmlCaptureOffice365DLConfiguration = get-date

    Out-LogFile -string "Capture the original office 365 distribution list information."

    $office365DLConfiguration=Get-O365DLConfiguration -groupSMTPAddress $groupSMTPAddress -errorAction STOP
    
    $office365GroupConfiguration = get-o365GroupConfiguration -groupSMTPAddress $groupSMTPAddress -errorAction STOP

    Out-LogFile -string $office365DLConfiguration

    Out-LogFile -string "Create an XML file backup of the office 365 DL configuration."

    Out-XMLFile -itemToExport $office365DLConfiguration -itemNameToExport $xmlFiles.office365DLConfigurationXML.value

    out-logfile -string $office365GroupConfiguration

    out-logfile -string "Create an XML file backup of the office 365 group cofniguration."

    out-xmlfile -itemToExport $office365GroupConfiguration -itemNameToExport $xmlFiles.office365GroupConfigurationXML.value

    $htmlCaptureGraphDLConfiguration = get-date

    $msGraphDLConfiguration = get-msGraphDLConfiguration -office365DLConfiguration $office365DLConfiguration -msGraphURL $msGraphURL -errorAction STOP

    out-logfile -string "Create an XML file backup of the Azure AD DL Configuration"

    out-xmlFile -itemToExport $msGraphDLConfiguration -itemNameToExport $xmlFiles.msGraphDLConfigurationXML.value

    $htmlCaptureGraphDLMembership = get-date

    $msGraphDLMembership = @(get-msGraphMembership -groupobjectID $msGraphDLConfiguration.id -msGraphURL $msGraphURL -errorAction STOP)

    if ($msGraphDLMembership.count -gt 0)
    {
            out-xmlFile -itemToExport $msGraphDLMembership -itemNameToExport $xmlFiles.msGraphDLMembershipXML.value
    }

    $htmlCaptureOffice365DLMembership = get-date

    $office365DLMembership = @(get-O365DLMembership -groupSMTPAddress $office365DLConfiguration.externalDirectoryObjectID -errorAction STOP)

    if ($office365DLMembership.count -gt 0)
    {
        out-logfile -string "Creating an XML file backup of the Office 365 Original DL Membership"

        out-xmlFile -itemToExport $office365DLMembership -itemNameToExport $xmlFiles.office365DLMembership.value
    }

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END GET ORIGINAL DL CONFIGURATION LOCAL AND CLOUD"
    Out-LogFile -string "********************************************************************************"

    $htmlStartGroupValidation = get-date

    Invoke-Office365SafetyCheck -o365dlconfiguration $office365DLConfiguration -azureADDLConfiguration $msGraphDLConfiguration -errorAction STOP

    $htmlStartAttributeNormalization = get-date
    $FunctionStartTime = get-universalDateTime

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN NORMALIZE DNS FOR ALL ATTRIBUTES"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the members DN to Office 365 identifier."

    if ($originalDLConfiguration.($onPremADAttributes.onPremMembers.Value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremMembers.Value))
        {
            #Resetting error variable.

            $isTestError="No"

            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -isMember:$TRUE -activeDirectoryAttribute $onPremADAttributes.onPremMembers.Value -activeDirectoryAttributeCommon $onPremADAttributes.onPremMembersCommon.Value -groupSMTPAddress $groupSMTPAddress -skipNestedGroupCheck $skipNestedGroupCheck -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else 
                {
                    $exchangeDLMembershipSMTP.add($normalizedTest)
                }
                
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeDLMembershipSMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the group:"
        
        out-logfile -string $exchangeDLMembershipSMTP
    }
    else 
    {
        out-logFile -string "The distribution group has no members."    
    }

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the reject members DN to Office 365 identifier."

    Out-LogFile -string "REJECT USERS"

    if ($originalDLConfiguration.($onPremADAttributes.onPremRejectMessagesFromSenders.value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremRejectMessagesFromSenders.value))
        {
            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremRejectMessagesFromSenders.value -activeDirectoryAttributeCommon $onPremADAttributes.onPremRejectMessagesFromSendersCommon.value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else 
                {
                    $exchangeRejectMessagesSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    Out-LogFile -string "REJECT GROUPS"

    if ($originalDLConfiguration.($onPremADAttributes.onPremRejectMessagesFromDLMembers.value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremRejectMessagesFromDLMembers.value))
        {
            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremRejectMessagesFromDLMembers.value -activeDirectoryAttributeCommon $onPremADAttributes.onPremRejectMessagesFromDLMembersCommon.value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else {
                    $exchangeRejectMessagesSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeRejectMessagesSMTP -ne $NULL)
    {
        out-logfile -string "The group has reject messages members."
        Out-logFile -string $exchangeRejectMessagesSMTP
    }
    else 
    {
        out-logfile "The group to be migrated has no reject messages from members."    
    }
    
    Out-LogFile -string "Invoke get-NormalizedDN to normalize the accept members DN to Office 365 identifier."

    Out-LogFile -string "ACCEPT USERS"

    if ($originalDLConfiguration.($onPremADAttributes.onPremAcceptMessagesFromSenders.value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremAcceptMessagesFromSenders.value))
        {
            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremAcceptMessagesFromSenders.value -activeDirectoryAttributeCommon $onPremADAttributes.onPremAcceptMessagesFromSendersCommon.value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else {
                    $exchangeAcceptMessagesSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logFile -string $_ -isError:$TRUE
            }
        }
    }

    Out-LogFile -string "ACCEPT GROUPS"

    if ($originalDLConfiguration.($onPremADAttributes.onPremAcceptMessagesFromDLMembers.value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremAcceptMessagesFromDLMembers.value))
        {
            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremAcceptMessagesFromDLMembers.value -activeDirectoryAttributeCommon $onPremADAttributes.onPremAcceptMessagesFromDLMembersCommon.value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else 
                {
                    $exchangeAcceptMessagesSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeAcceptMessagesSMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the accept messages from senders:"
        
        out-logfile -string $exchangeAcceptMessagesSMTP
    }
    else
    {
        out-logFile -string "This group has no accept message from restrictions."    
    }
    
    Out-LogFile -string "Invoke get-NormalizedDN to normalize the managedBy members DN to Office 365 identifier."

    Out-LogFile -string "Process MANAGEDBY"

    if ($originalDLConfiguration.($onPremADAttributes.onPremManagedBy.Value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremManagedBy.Value))
        {
            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremManagedBy.Value -activeDirectoryAttributeCommon $onPremADAttributes.onPremManagedByCommon.Value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else 
                {
                    $exchangeManagedBySMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    Out-LogFile -string "Process CoMANAGERS"

    if ($originalDLConfiguration.($onPremADAttributes.onPremCoManagedBy.Value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremCoManagedBy.Value))
        {
            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremCoManagedBy.Value -activeDirectoryAttributeCommon $onPremADAttributes.onPremCoManagedByCommon.Value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else 
                {
                    $exchangeManagedBySMTP+=$normalizedTest
                }
                
            }
            catch 
            {
                out-logFile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeManagedBySMTP -ne $NULL)
    {
        #First scan is to ensure that any of the groups listed on the managed by objects are still security.
        #It is possible someone added it to managed by and changed the group type after.

        foreach ($object in $exchangeManagedBySMTP)
        {
            #If the objec thas a non-null group type (is a group) and the value of the group type matches none of the secuity group types.
            #The object is a distribution list - no good.

            if (($object.groupType -ne $NULL) -and ($object.groupType -ne "-2147483640") -and ($object.groupType -ne "-2147483646") -and ($object.groupType -ne "-2147483644"))
            {
                $object.isError=$TRUE
                $object.isErrorMessage = "GROUP_NO_LONGER_SECURITY_EXCEPTION: A group was found on the owners attribute that is no longer a security group.  Security group is required.  Remove group or change group type to security."
                
                out-logfile -string object

                $global:preCreateErrors+=$object

                out-logfile -string "A distribution list (not security enabled) was found on managed by."
                out-logfile -string "The group must be converted to security or removed from managed by."
                out-logfile -string $object.primarySMTPAddressOrUPN
            }
        }

        Out-LogFile -string "The following objects are members of the managedBY:"
        
        out-logfile -string $exchangeManagedBySMTP
    }
    else 
    {
        out-logfile -string "The group has no managers."    
    }

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the moderatedBy members DN to Office 365 identifier."

    Out-LogFile -string "Process MODERATEDBY"

    if ($originalDLConfiguration.($onPremADAttributes.onPremModeratedBy.Value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremModeratedBy.Value))
        {
            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremModeratedBy.Value -activeDirectoryAttributeCommon $onPremADAttributes.onPremModeratedByCommon.Value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else 
                {
                    $exchangeModeratedBySMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeModeratedBySMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the moderatedBY:"
        
        out-logfile -string $exchangeModeratedBySMTP    
    }
    else 
    {
        out-logfile "The group has no moderators."    
    }

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the bypass moderation users members DN to Office 365 identifier."

    Out-LogFile -string "Process BYPASS USERS"

    if ($originalDLConfiguration.($onPremADAttributes.onPremBypassModerationFromSenders.Value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremBypassModerationFromSenders.Value))
        {
            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremBypassModerationFromSenders.Value -activeDirectoryAttributeCommon $onPremADAttributes.onPremBypassModerationFromSendersCommon.Value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else 
                {
                    $exchangeBypassModerationSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logFile -string $_ -isError:$TRUE
            }
        }
    }

    Out-LogFile -string "Invoke get-NormalizedDN to normalize the bypass moderation groups members DN to Office 365 identifier."

    Out-LogFile -string "Process BYPASS GROUPS"

    if ($originalDLConfiguration.($onPremADAttributes.onPremBypassModerationFromDL.Value) -ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremBypassModerationFromDL.Value))
        {
            try 
            {
                $normalizedTest = get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremBypassModerationFromDL.Value -activeDirectoryAttributeCommon $onPremADAttributes.onPremBypassModerationFromDLCommon.Value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else 
                {
                    $exchangeBypassModerationSMTP+=$normalizedTest
                }
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeBypassModerationSMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the bypass moderation:"
        
        out-logfile -string $exchangeBypassModerationSMTP 
    }
    else 
    {
        out-logfile "The group has no bypass moderation."    
    }

    if ($originalDLConfiguration.($onPremADAttributes.onPremGrantSendOnBehalfTo.Value)-ne $NULL)
    {
        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremGrantSendOnBehalfTo.Value))
        {
            try 
            {
                $normalizedTest=get-normalizedDN -globalCatalogServer $corevariables.globalCatalogWithPort.value -DN $DN -adCredential $activeDirectoryCredential -originalGroupDN $originalDLConfiguration.distinguishedName -activeDirectoryAttribute $onPremADAttributes.onPremGrantSendOnBehalfTo.Value -activeDirectoryAttributeCommon $onPremADAttributes.onPremGrantSendOnBehalfToCommon.Value -groupSMTPAddress $groupSMTPAddress -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP -cn "None"

                out-logfile -string $normalizedTest

                if ($normalizedTest.isError -eq $TRUE)
                {
                    $global:preCreateErrors+=$normalizedTest
                }
                else 
                {
                    $exchangeGrantSendOnBehalfToSMTP+=$normalizedTest
                }
                
            }
            catch 
            {
                out-logfile -string $_ -isError:$TRUE
            }
        }
    }

    if ($exchangeGrantSendOnBehalfToSMTP -ne $NULL)
    {
        Out-LogFile -string "The following objects are members of the grant send on behalf to:"
        
        out-logfile -string $exchangeGrantSendOnBehalfToSMTP
    }
    else 
    {
        out-logfile "The group has no grant send on behalf to."    
    }

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END NORMALIZE DNS FOR ALL ATTRIBUTES"
    Out-LogFile -string "********************************************************************************"

    $FunctionEndTime = get-universalDateTime

    $telemetryNormalizeDN = get-elapsedTime -startTime $FunctionStartTime -endTime $FunctionEndTime

    out-logfile -string ("Time to Normalize DNs: "+$telemetryNormalizeDN.toString())

    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"
    out-logFile -string "Summary of group information:"
    out-logfile -string ("The number of objects included in the member migration: "+$exchangeDLMembershipSMTP.count)
    out-logfile -string ("The number of objects included in the reject memebers: "+$exchangeRejectMessagesSMTP.count)
    out-logfile -string ("The number of objects included in the accept memebers: "+$exchangeAcceptMessagesSMTP.count)
    out-logfile -string ("The number of objects included in the managedBY memebers: "+$exchangeManagedBySMTP.count)
    out-logfile -string ("The number of objects included in the moderatedBY memebers: "+$exchangeModeratedBySMTP.count)
    out-logfile -string ("The number of objects included in the bypassModeration memebers: "+$exchangeBypassModerationSMTP.count)
    out-logfile -string ("The number of objects included in the grantSendOnBehalfTo memebers: "+$exchangeGrantSendOnBehalfToSMTP.count)
    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"

    out-logfile -string "If a group was migrated by DLConversionV2 it is possible it has special case objects on it that served the migration."
    out-logfile -string "Test each of the normalized arrays - if any of those were located recommend migration with DLConversionV2 so that all dependencies are handled in the migration."

    $global:dlConversionV2Test=@()
    $global:dlConversionV2Test+= @($exchangeDLMembershipSMTP | where {$_.isAlreadyMigrated -eq $true })
    $global:dlConversionV2Test+= @($exchangeRejectMessagesSMTP | where {$_.isAlreadyMigrated -eq $true })
    $global:dlConversionV2Test+= @($exchangeAcceptMessagesSMTP | where {$_.isAlreadyMigrated -eq $true })
    $global:dlConversionV2Test+= @($exchangeManagedBySMTP | where {$_.isAlreadyMigrated -eq $true })
    $global:dlConversionV2Test+= @($exchangeModeratedBySMTP | where {$_.isAlreadyMigrated -eq $true })
    $global:dlConversionV2Test+= @($exchangeBypassModerationSMTP | where {$_.isAlreadyMigrated -eq $true })
    $global:dlConversionV2Test+= @($exchangeGrantSendOnBehalfToSMTP | where {$_.isAlreadyMigrated -eq $true })

    if (($global:dlConversionV2Test.count -gt 0) -and ($isHealthCheck -eq $false))
    {
        generate-HTMLFile

        start-sleep -s 10

        out-logfile -string "Error - members or properties of this DL have dependencies on DLConversionV2 migration."
        out-logfile -string "Recommend the DL be migrated with DLConversionV2"
        out-logfile -string "ERROR_DLCONVERSIONV2_RECOMMENDED_EXCEPTION" -isError:$TRUE
    }


    $htmlStartCloudValidation = get-date
    $htmlStartCloudValidationOnPremises = get-Date

    $FunctionStartTime = get-universalDateTime

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN VALIDATE RECIPIENTS IN CLOUD"
    Out-LogFile -string "********************************************************************************"

    out-logfile -string "Begin accepted domain validation."

    test-AcceptedDomain -originalDLConfiguration $originalDlConfiguration -errorAction STOP

    out-logfile -string "Test for centralized mail transport."

    test-outboundConnector -overrideCentralizedMailTransportEnabled $overrideCentralizedMailTransportEnabled -errorAction STOP

    if ($customRoutingDomain -eq "")
    {
        out-logfile -string "Determine the mail onmicrosoft domain necessary for cross premises routing."
        try {
            $mailOnMicrosoftComDomain = Get-MailOnMicrosoftComDomain -errorAction STOP
        }
        catch {
            out-logfile -string $_
            out-logfile -string "Unable to obtain the onmicrosoft.com domain." -errorAction STOP    
        }
    }
    else 
    {
        out-logfile -string "The administrtor has specified a custome routing domain - maybe for legacy tenant implementations."

        $mailOnMicrosoftComDomain = $customRoutingDomain
    }

    if ($testRecipientHealth -eq $TRUE)
    {
        out-logfile -string "Being validating all distribution list members."
    
        if ($exchangeDLMembershipSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL member is in Office 365 / Exchange Online"

            foreach ($member in $exchangeDLMembershipSMTP)
            {
                #Reset the failure.

                $isTestError="No"

                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-O365Recipient -member $member

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_DEPENDENCY_NOT_FOUND_EXCEPTION: A group dependency was not found in Office 365.  Please either ensure the dependency is present or remove the dependency from the group."

                        out-logfile -string $member

                        $global:testOffice365Errors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There are no DL members to test."    
        }

        out-logfile -string "Begin evaluating all members with reject rights."

        if ($exchangeRejectMessagesSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL reject messages is in Office 365."

            foreach ($member in $exchangeRejectMessagesSMTP)
            {
                #Reset error variable.

                $isTestError="No"

                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-O365Recipient -member $member

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_DEPENDENCY_NOT_FOUND_EXCEPTION: A group dependency was not found in Office 365.  Please either ensure the dependency is present or remove the dependency from the group."

                        out-logfile -string $member

                        $global:testOffice365Errors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There are no reject members to test."    
        }

        out-logfile -string "Begin evaluating all members with accept rights."

        if ($exchangeAcceptMessagesSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL accept messages is in Office 365 / Exchange Online"

            foreach ($member in $exchangeAcceptMessagesSMTP)
            {
                #Reset error variable.

                $isTestError="No"
                
                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-O365Recipient -member $member

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_DEPENDENCY_NOT_FOUND_EXCEPTION: A group dependency was not found in Office 365.  Please either ensure the dependency is present or remove the dependency from the group."

                        out-logfile -string $member

                        $global:testOffice365Errors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There are no accept members to test."    
        }

        out-logfile -string "Begin evaluating all managed by members."

        if ($exchangeManagedBySMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL managed by is in Office 365 / Exchange Online"

            foreach ($member in $exchangeManagedBySMTP)
            {
                #Reset Error Variable.

                $isTestError="No"
                
               out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-O365Recipient -member $member

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_DEPENDENCY_NOT_FOUND_EXCEPTION: A group dependency was not found in Office 365.  Please either ensure the dependency is present or remove the dependency from the group."

                        out-logfile -string $member

                        $global:testOffice365Errors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There were no managed by members to evaluate."    
        }

        out-logfile -string "Begin evaluating all moderated by members."

        if ($exchangeModeratedBySMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL moderated by is in Office 365 / Exchange Online"

            foreach ($member in $exchangeModeratedBySMTP)
            {
                #Reset error variable.

                $isTestError="No"

                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-O365Recipient -member $member

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_DEPENDENCY_NOT_FOUND_EXCEPTION: A group dependency was not found in Office 365.  Please either ensure the dependency is present or remove the dependency from the group."

                        out-logfile -string $member

                        $global:testOffice365Errors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There were no moderated by members to evaluate."    
        }

        out-logfile -string "Being evaluating all bypass moderation members."

        if ($exchangeBypassModerationSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL bypass moderation is in Office 365 / Exchange Online"

            foreach ($member in $exchangeBypassModerationSMTP)
            {
                #Reset error variable.

                $isTestError="No"

               out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-O365Recipient -member $member

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_DEPENDENCY_NOT_FOUND_EXCEPTION: A group dependency was not found in Office 365.  Please either ensure the dependency is present or remove the dependency from the group."

                        out-logfile -string $member

                        $global:testOffice365Errors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There were no bypass moderation members to evaluate."    
        }

        out-logfile -string "Begin evaluation of all grant send on behalf to members."

        if ($exchangeGrantSendOnBehalfToSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL grant send on behalf to is in Office 365 / Exchange Online"

            foreach ($member in $exchangeGrantSendOnBehalfToSMTP)
            {
                $isTestError = "No"

                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-O365Recipient -member $member

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_DEPENDENCY_NOT_FOUND_EXCEPTION: A group dependency was not found in Office 365.  Please either ensure the dependency is present or remove the dependency from the group."

                        out-logfile -string $member

                        $global:testOffice365Errors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There were no grant send on behalf to members to evaluate."    
        }
    }

    $htmlStartCloudValidationOffice365 = get-Date

    out-logfile -string "Determine if individual properties should be reviewed to ensure they match."

    if ($testPropertyHealth -eq $TRUE)
    {
        out-logfile -string "Being validating all distribution list members."
    
        if ($exchangeDLMembershipSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL member is in Office 365 / Exchange Online"

            foreach ($member in $exchangeDLMembershipSMTP)
            {
                #Reset the failure.

                $isTestError="No"

                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-O365Member -member $member -membership $office365DLMembership

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_MEMBER_NOT_FOUND_IN_GROUP_EXCEPTION: An on premises group member was not found as a member of the Office 365 Distribution List."

                        out-logfile -string $member

                        $global:testOffice365PropertyErrors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There are no DL members to test."    
        }

        out-logfile -string "Begin evaluating all members with reject rights."

        if ($exchangeRejectMessagesSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL reject messages is in Office 365."

            foreach ($member in $exchangeRejectMessagesSMTP)
            {
                #Reset error variable.

                $isTestError="No"

                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-o365Property -member $member -membership $office365DLConfiguration.RejectMessagesFromSendersOrMembers

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_MEMBER_NOT_FOUND_EXCEPTION: An on-premsies recipient with RejectMessagesFromSendersOrMembers rights not found on Office 365 DL."

                        out-logfile -string $member

                        $global:testOffice365PropertyErrors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There are no reject members to test."    
        }

        out-logfile -string "Begin evaluating all members with accept rights."

        if ($exchangeAcceptMessagesSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL accept messages is in Office 365 / Exchange Online"

            foreach ($member in $exchangeAcceptMessagesSMTP)
            {
                #Reset error variable.

                $isTestError="No"
                
                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-o365Property -member $member -membership $office365DLConfiguration.AcceptMessagesOnlyFromSendersOrMembers

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_MEMBER_NOT_FOUND_EXCEPTION: An on-premsies recipient with AcceptMessagesFromSendersOrMembers rights not found on Office 365 DL."

                        out-logfile -string $member

                        $global:testOffice365PropertyErrors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There are no accept members to test."    
        }

        out-logfile -string "Begin evaluating all managed by members."

        if ($exchangeManagedBySMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL managed by is in Office 365 / Exchange Online"

            foreach ($member in $exchangeManagedBySMTP)
            {
                #Reset Error Variable.

                $isTestError="No"
                
               out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-o365Property -member $member -membership $office365DLConfiguration.ManagedBy

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_MEMBER_NOT_FOUND_EXCEPTION: An on-premsies recipient with ManagedBy rights not found on Office 365 DL."

                        out-logfile -string $member

                        $global:testOffice365PropertyErrors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There were no managed by members to evaluate."    
        }

        out-logfile -string "Begin evaluating all moderated by members."

        if ($exchangeModeratedBySMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL moderated by is in Office 365 / Exchange Online"

            foreach ($member in $exchangeModeratedBySMTP)
            {
                #Reset error variable.

                $isTestError="No"

                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-o365Property -member $member -membership $office365DLConfiguration.ModeratedBy

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_MEMBER_NOT_FOUND_EXCEPTION: An on-premsies recipient with ModeratedBy rights not found on Office 365 DL."

                        out-logfile -string $member

                        $global:testOffice365PropertyErrors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There were no moderated by members to evaluate."    
        }

        out-logfile -string "Being evaluating all bypass moderation members."

        if ($exchangeBypassModerationSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL bypass moderation is in Office 365 / Exchange Online"

            foreach ($member in $exchangeBypassModerationSMTP)
            {
                #Reset error variable.

                $isTestError="No"

               out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-o365Property -member $member -membership $office365DLConfiguration.BypassModerationFromSendersOrMembers

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_MEMBER_NOT_FOUND_EXCEPTION: An on-premsies recipient with BypassModerationFromSendersOrMembers rights not found on Office 365 DL."

                        out-logfile -string $member

                        $global:testOffice365PropertyErrors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There were no bypass moderation members to evaluate."    
        }

        out-logfile -string "Begin evaluation of all grant send on behalf to members."

        if ($exchangeGrantSendOnBehalfToSMTP.count -gt 0)
        {
            out-logfile -string "Ensuring each DL grant send on behalf to is in Office 365 / Exchange Online"

            foreach ($member in $exchangeGrantSendOnBehalfToSMTP)
            {
                $isTestError = "No"

                out-LogFile -string ("Testing = "+$member.primarySMTPAddressOrUPN)

                try{
                    $isTestError=test-o365Property -member $member -membership $office365DLConfiguration.GrantSendOnBehalfTo

                    if ($isTestError -eq "Yes")
                    {
                        $member.isError = $TRUE
                        $member.isErrorMessage = "OFFICE_365_MEMBER_NOT_FOUND_EXCEPTION: An on-premsies recipient with GrantSendOnBehalfTo rights not found on Office 365 DL."

                        out-logfile -string $member

                        $global:testOffice365PropertyErrors += $member | ConvertTo-JSON | ConvertFrom-JSON
                    }
                }
                catch{
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }
        else 
        {
            out-logfile -string "There were no grant send on behalf to members to evaluate."    
        }
    }

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END VALIDATE RECIPIENTS IN CLOUD"
    Out-LogFile -string "********************************************************************************"

    $FunctionEndTime = get-universalDateTime

    $telemetryValidateCloudRecipients = get-elapsedTime -startTime $FunctionStartTime -endTime $FunctionEndTime

    out-logfile -string ("Time to validate recipients in cloud: "+ $telemetryValidateCloudRecipients.toString())

    if (($global:preCreateErrors.count -gt 0) -or ($global:testOffice365Errors.count -gt 0) -or ($global:testOffice365PropertyErrors.count -gt 0))
    {
        #Write the XML files first so that the error table is complete without separation.

        if ($global:preCreateErrors.count -gt 0)
        {
            out-xmlFile -itemToExport $global:preCreateErrors -itemNameToExport $xmlFiles.preCreateErrorsXML.value
        }

        if ($global:testOffice365Errors.Count -gt 0)
        {
            out-xmlFile -itemToExport $global:testOffice365Errors -itemNametoExport $xmlfiles.testOffice365ErrorsXML.value
        }

        if ($global:testOffice365PropertyErrors.count -gt 0)
        {
            out-xmlFile -itemToExport $global:testOffice365PropertyErrors -itemNametoExport $xmlfiles.testOffice365PropertyErrorsXML.value
        }

        out-logfile -string "+++++"
        out-logfile -string "Pre-requist checks failed.  Please refer to the following list of items that require addressing for migration to proceed."
        out-logfile -string "+++++"
        out-logfile -string ""

        if ($global:preCreateErrors.count -gt 0)
        {
            foreach ($preReq in $global:preCreateErrors)
            {
                write-errorEntry -errorEntry $preReq
            }
        }

        if ($global:testOffice365Errors.count -gt 0)
        {
            foreach ($preReq in $global:testOffice365Errors)
            {
                write-errorEntry -errorEntry $prereq
            }
        }

        if ($global:testOffice365PropertyErrors.count -gt 0)
        {
            foreach ($preReq in $global:testOffice365PropertyErrors)
            {
                write-errorEntry -errorEntry $preReq
            }
        }

        if ($isHealthCheck -eq $FALSE)
        {
            generate-HTMLFile

            start-sleep -s 5

            out-logfile -string "Pre-requiste checks failed.  Please refer to the previous list of items that require addressing for migration to proceed." -isError:$TRUE
        }
        else
        {
            out-logfile -string "Pre-requiste checks failed.  Please refer to the previous list of items that require addressing for migration to proceed."
        }  
    }

    if ($isHealthCheck -eq $TRUE)
    {
        return
    }

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "BEGIN RECORD DEPENDENCIES ON MIGRATED GROUP"
    Out-LogFile -string "********************************************************************************"

    $htmlCaptureOnPremisesDependencies = get-date

    $FunctionStartTime = get-universalDateTime

    out-logfile -string "Get all the groups that this user is a member of - normalize to canonicalname."

    #Start with groups this DL is a member of remaining on premises.

    if ($originalDLConfiguration.($onPremADAttributes.onPremMemberOf.value) -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremMemberOf.value))
        {
            $allGroupsMemberOf += get-canonicalname -globalCatalog $corevariables.globalCatalogWithPort.value -dn $DN -adCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP
        }
    }

    #Handle all recipients that have forwarding to this group based on forwarding address.

    if ($originalDLConfiguration.($onPremADAttributes.onPremForwardingAddressBL.value) -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremForwardingAddressBL.value))
        {
            $allUsersForwardingAddress += get-canonicalname -globalCatalog $corevariables.globalCatalogWithPort.value -dn $DN -adCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP
        }
    }

    #Handle all groups this object has reject permissions on.

    if ($originalDLConfiguration.($onPremADAttributes.onPremRejectMessagesFromDLMembersBL.value) -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremRejectMessagesFromDLMembersBL.value))
        {
            $allGroupsReject += get-canonicalname -globalCatalog $corevariables.globalCatalogWithPort.value -dn $DN -adCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP
        }
    }

    #Handle all groups this object has accept permissions on.

    if ($originalDLConfiguration.($onPremADAttributes.onPremAcceptMessagesFromDLMembersBL.value) -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremAcceptMessagesFromDLMembersBL.value))
        {
            $allGroupsAccept += get-canonicalname -globalCatalog $corevariables.globalCatalogWithPort.value -dn $DN -adCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP
        }
    }

    if ($originalDlConfiguration.($onPremADAttributes.onPremCoManagedByBL.value) -ne $NULL)
    {
        out-logfile -string "Calling get canonical name."

        foreach ($dn in $originalDLConfiguration.($onPremADAttributes.onPremCoManagedByBL.value))
        {
            $allGroupsCoManagedByBL += get-canonicalName -globalCatalog $corevariables.globalCatalogWithPort.value -dn $DN -adCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP
        }
    }
    else 
    {
        out-logfile -string "The group is not a co manager on any other groups."    
    }

    #Handle all groups this object has bypass moderation permissions on.

    if ($originalDLConfiguration.($onPremADAttributes.onPremBypassModerationFromDLMembersBL.value) -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremBypassModerationFromDLMembersBL.value))
        {
            $allGroupsBypassModeration += get-canonicalname -globalCatalog $corevariables.globalCatalogWithPort.value -dn $DN -adCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP
        }
    }

    #Handle all groups this object has accept permissions on.

    if ($originalDLConfiguration.($onPremADAttributes.onPremGrantSendOnBehalfToBL.value) -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremGrantSendOnBehalfToBL.value))
        {
            $allGroupsGrantSendOnBehalfTo += get-canonicalname -globalCatalog $corevariables.globalCatalogWithPort.value -dn $DN -adCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP
        }
    }

    #Handle all groups this object has manager permissions on.

    if ($originalDLConfiguration.($onPremADAttributes.onPremCoManagedByBL.value) -ne $NULL)
    {
        out-logfile -string "Calling get-CanonicalName."

        foreach ($DN in $originalDLConfiguration.($onPremADAttributes.onPremCoManagedByBL.value))
        {
            $allGroupsManagedBy += get-canonicalname -globalCatalog $corevariables.globalCatalogWithPort.value -dn $DN -adCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -errorAction STOP
        }
    }

    $FunctionEndTime = get-universalDateTime

    $telemetryDependencyOnPrem = get-elapsedTime -startTime $FunctionStartTime -endTime $FunctionEndTime

    out-logfile -string ("Time to calculate on premsies dependencies: "+ $telemetryDependencyOnPrem.toString())

    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"
    out-logfile -string ("Summary of dependencies found:")
    out-logfile -string ("The number of groups that the migrated DL is a member of = "+$allGroupsMemberOf.count)
    out-logfile -string ("The number of groups that this group is a manager of: = "+$allGroupsManagedBy.count)
    out-logfile -string ("The number of groups that this group has grant send on behalf to = "+$allGroupsGrantSendOnBehalfTo.count)
    out-logfile -string ("The number of groups that have this group as bypass moderation = "+$allGroupsBypassModeration.count)
    out-logfile -string ("The number of groups with accept permissions = "+$allGroupsAccept.count)
    out-logfile -string ("The number of groups with reject permissions = "+$allGroupsReject.count)
    out-logfile -string ("The number of mailboxes forwarding to this group is = "+$allUsersForwardingAddress.count)
    out-logfile -string ("The number of groups this group is a co-manager on = "+$allGroupsCoManagedByBL.Count)
    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"

    #Exit #Debug exit

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END RECORD DEPENDENCIES ON MIGRATED GROUP"
    Out-LogFile -string "********************************************************************************"

    Out-LogFile -string "Recording all gathered information to XML to preserve original values."
    
    if ($exchangeDLMembershipSMTP -ne $NULL)
    {
        Out-XMLFile -itemtoexport $exchangeDLMembershipSMTP -itemNameToExport $xmlFiles.exchangeDLMembershipSMTPXML.value
    }

    if ($exchangeRejectMessagesSMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeRejectMessagesSMTP -itemNameToExport $xmlFiles.exchangeRejectMessagesSMTPXML.value
    }

    if ($exchangeAcceptMessagesSMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeAcceptMessagesSMTP -itemNameToExport $xmlFiles.exchangeAcceptMessagesSMTPXML.value
    }

    if ($exchangeManagedBySMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeManagedBySMTP -itemNameToExport $xmlFiles.exchangeManagedBySMTPXML.value
    }

    if ($exchangeModeratedBySMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeModeratedBySMTP -itemNameToExport $xmlFiles.exchangeModeratedBySMTPXML.value
    }

    if ($exchangeBypassModerationSMTP -ne $NULL)
    {
        out-xmlfile -itemtoexport $exchangeBypassModerationSMTP -itemNameToExport $xmlFiles.exchangeBypassModerationSMTPXML.value
    }

    if ($exchangeGrantSendOnBehalfToSMTP -ne $NULL)
    {
        out-xmlfile -itemToExport $exchangeGrantSendOnBehalfToSMTP -itemNameToExport $xmlFiles.exchangeGrantSendOnBehalfToSMTPXML.value
    }

    if ($allGroupsMemberOf -ne $NULL)
    {
        out-xmlfile -itemtoexport $allGroupsMemberOf -itemNameToExport $xmlFiles.allGroupsMemberOfXML.value
    }
    
    if ($allGroupsReject -ne $NULL)
    {
        out-xmlfile -itemtoexport $allGroupsReject -itemNameToExport $xmlFiles.allGroupsRejectXML.value
    }
    
    if ($allGroupsAccept -ne $NULL)
    {
        out-xmlfile -itemtoexport $allGroupsAccept -itemNameToExport $xmlFiles.allGroupsAcceptXML.value
    }

    if ($allGroupsCoManagedByBL -ne $NULL)
    {
        out-xmlfile -itemToExport $allGroupsCoManagedByBL -itemNameToExport $xmlFiles.allGroupsCoManagedByXML.value
    }

    if ($allGroupsBypassModeration -ne $NULL)
    {
        out-xmlfile -itemtoexport $allGroupsBypassModeration -itemNameToExport $xmlFiles.allGroupsBypassModerationXML.value
    }

    if ($allUsersForwardingAddress -ne $NULL)
    {
        out-xmlFile -itemToExport $allUsersForwardingAddress -itemNameToExport $xmlFiles.allUsersForwardingAddressXML.value
    }

    if ($allGroupsManagedBy -ne $NULL)
    {
        out-xmlFile -itemToExport $allGroupsManagedBy -itemNameToExport $xmlFiles.allGroupsManagedByXML.value
    }

    if ($allGroupsGrantSendOnBehalfTo -ne $NULL)
    {
        out-xmlFile -itemToExport $allGroupsGrantSendOnBehalfTo -itemNameToExport $xmlFiles.allGroupsGrantSendOnBehalfToXML.value
    }

    #Ok so at this point we have preserved all of the information regarding the on premises DL.
    #It is possible that there could be cloud only objects that this group was made dependent on.
    #For example - the dirSync group could have been added as a member of a cloud only group - or another group that was migrated.
    #The issue here is that this gets VERY expensive to track - since some of the word to do do is not filterable.
    #With the LDAP improvements we no longer offert the option to track on premises - but the administrator can choose to track the cloud

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "START RETAIN OFFICE 365 GROUP DEPENDENCIES"
    Out-LogFile -string "********************************************************************************"

    $htmlRecordOffice365Dependencies = get-date

    $telemetryFunctionStartTime = get-universalDateTime

    out-logfile -string "Obtain all office 365 member of."
    $allOffice365MemberOf = Get-O365GroupDependency -dn $office365DLConfiguration.distinguishedName -attributeType $office365Attributes.office365Members.value -errorAction STOP
    out-logfile -string "Obtain all office 365 accept messages from senders or members."
    $allOffice365Accept = Get-O365GroupDependency -dn $office365DLConfiguration.distinguishedName -attributeType $office365Attributes.office365AcceptMessagesFrom.value -errorAction STOP
    out-logfile -string "Obtain all office 365 reject messages from senders or members."
    $allOffice365Reject = Get-O365GroupDependency -dn $office365DLConfiguration.distinguishedName -attributeType $office365Attributes.office365RejectMessagesFrom.value -errorAction STOP
    out-logfile -string "Obtain all office 365 bypass moderation from senders or members."
    $allOffice365BypassModeration = Get-O365GroupDependency -dn $office365DLConfiguration.distinguishedName -attributeType $office365Attributes.office365BypassModerationFrom.value -errorAction STOP
    out-logfile -string "Obtain all office 365 grant send on behalf to."
    $allOffice365GrantSendOnBehalfTo = Get-O365GroupDependency -dn $office365DLConfiguration.distinguishedName -attributeType $office365Attributes.office365GrantSendOnBehalfTo.value -errorAction STOP
    out-logfile -string "Obtain all office 365 managedBy."
    $allOffice365ManagedBy = Get-O365GroupDependency -dn $office365DLConfiguration.distinguishedName -attributeType $office365Attributes.office365ManagedBy.value -errorAction STOP
    out-logfile -string "Obtain all office 365 forwardingAddress."
    $allOffice365ForwardingAddress = Get-O365GroupDependency -dn $office365DLConfiguration.distinguishedName -attributeType $office365Attributes.office365ForwardingAddress.value -errorAction STOP
    out-logfile -string "Obtain all office 365 Send As on Others."
    $allOffice365SendAsAccess = Get-O365DLSendAs -groupSMTPAddress $office365DLConfiguration.externalDirectoryObjectID -isTrustee:$TRUE -office365GroupConfiguration $office365GroupConfiguration -errorAction STOP
    out-logfile -string "Obtain all office 365 Send As on Group."
    $allOffice365SendAsAccessOnGroup = get-o365DLSendAs -groupSMTPAddress $office365DLConfiguration.externalDirectoryObjectID -errorAction STOP

    if ($allOffice365MemberOf -ne $NULL)
    {
        out-xmlfile -itemtoexport $allOffice365MemberOf -itemNameToExport $xmlFiles.allOffice365MemberOfXML.value
    }

    if ($allOffice365Accept -ne $NULL)
    {
        out-xmlFile -itemToExport $allOffice365Accept -itemNameToExport $xmlFiles.allOffice365AcceptXML.value
    }

    if ($allOffice365Reject -ne $NULL)
    {
        out-xmlFile -itemToExport $allOffice365Reject -itemNameToExport $xmlFiles.allOffice365RejectXML.value
    }
    
    if ($allOffice365BypassModeration -ne $NULL)
    {
        out-xmlFile -itemToExport $allOffice365BypassModeration -itemNameToExport $xmlFiles.allOffice365BypassModerationXML.value
    }

    if ($allOffice365GrantSendOnBehalfTo -ne $NULL)
    {
        out-xmlfile -itemToExport $allOffice365GrantSendOnBehalfTo -itemNameToExport $xmlFiles.allOffice365GrantSendOnBehalfToXML.value
    }

    if ($allOffice365ManagedBy -ne $NULL)
    {
        out-xmlFile -itemToExport $allOffice365ManagedBy -itemNameToExport $xmlFiles.allOffice365ManagedByXML.value
    }

    if ($allOffice365ForwardingAddress -ne $NULL)
    {
        out-xmlfile -itemToExport $allOffice365ForwardingAddress -itemNameToExport $xmlFiles.allOffice365ForwardingAddressXML.value
    }

    if ($allOffice365SendAsAccess -ne $NULL)
    {
        out-xmlfile -itemToExport $allOffice365SendAsAccess -itemNameToExport $xmlFiles.allOffic365SendAsAccessXML.value
    }

    if ($allOffice365SendAsAccessOnGroup -ne $NULL)
    {
        out-xmlfile -itemToExport $allOffice365SendAsAccessOnGroup -itemNameToExport $xmlFiles.allOffice365SendAsAccessOnGroupXML.value
    }

    $telemetryFunctionEndTime = get-universalDateTime

    $telemetryCollectOffice365Dependency = ($telemetryFunctionEndTime - $telemetryFunctionStartTime).seconds

    out-logfile -string ("Time to gather Office 365 dependencies: "+$telemetryCollectOffice365Dependency.tostring())

    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"
    out-logfile -string ("Summary of dependencies found:")
    out-logfile -string ("The number of office 365 objects that the migrated DL is a member of = "+$allOffice365MemberOf.count)
    out-logfile -string ("The number of office 365 objects that this group is a manager of: = "+$allOffice365ManagedBy.count)
    out-logfile -string ("The number of office 365 objects that this group has grant send on behalf to = "+$allOffice365GrantSendOnBehalfTo.count)
    out-logfile -string ("The number of office 365 objects that have this group as bypass moderation = "+$allOffice365BypassModeration.count)
    out-logfile -string ("The number of office 365 objects with accept permissions = "+$allOffice365Accept.count)
    out-logfile -string ("The number of office 365 objects with reject permissions = "+$allOffice365Reject.count)
    out-logfile -string ("The number of office 365 mailboxes forwarding to this group is = "+$allOffice365ForwardingAddress.count)
    out-logfile -string ("The number of recipients that have send as rights on the group to be migrated = "+$allOffice365SendAsAccessOnGroup.count)
    out-logfile -string ("The number of office 365 recipients where the group has send as rights = "+$allOffice365SendAsAccess.count)
    out-logfile -string "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/"

    
    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END RETAIN OFFICE 365 GROUP DEPENDENCIES"
    Out-LogFile -string "********************************************************************************"

    $htmlSetGroupCloudOnly = Get-Date

    $telemetryFunctionStartTime = get-universalDateTime

    out-logfile -string "Attempt to set the group to cloud only status."

    set-DLCloudOnly -msGraphURL $msGraphURL -office365DLConfiguration $office365DLConfiguration 

    test-CloudDLPresentGraph -groupSMTPAddress $office365DLConfiguration.externalDirectoryObjectID -msGraphURL $msGraphURL -errorAction STOP

    $msGraphDLConfigurationPostMigration = get-msGraphDLConfiguration -office365DLConfiguration $office365DLConfiguration -msGraphURL $msGraphURL -errorAction STOP

    out-xmlFile -itemToExport $msGraphDLConfigurationPostMigration -itemNameToExport $xmlFiles.msGraphDLConfigurationPostMigrationXML.value

    $telemetryFunctionEndTime = get-universalDateTime

    $telemetryConvertGroupCloudOnly = get-elapsedTime -startTime $FunctionStartTime -endTime $FunctionEndTime

    $htmlTestExchangeOnlineCloudOnly = Get-Date

    $telemetryfunctionStartTime = get-universalDateTime

    test-CloudDLPresentExchangeOnline -groupSMTPAddress $office365DLConfiguration.externalDirectoryObjectID -errorAction STOP

    $telemetryfunctionEndTime = get-universalDateTime

    $telemetryConvertGroupCloudOnlyExchangeOnline = get-elapsedTime -startTime $telemetryfunctionStartTime -endTime $telemetryfunctionEndTime

    $htmlCaptureOffice365InfoPostMigration = Get-Date

    $office365DLConfigurationPostMigration = Get-O365DLConfiguration -groupSMTPAddress $office365DLConfiguration.externalDirectoryObjectID -errorAction STOP
    out-xmlFile -itemToExport $office365DLConfigurationPostMigration -itemNameToExport $xmlFiles.office365DLConfigurationPostMigrationXML.value

    $office365DLMembershipPostMigration = @(get-O365DLMembership -groupSMTPAddress $office365DLConfiguration.externalDirectoryObjectID -errorAction STOP)
    out-xmlFile -itemToExport $office365DLMembershipPostMigration -itemNametoExport $xmlFiles.office365DLMembershipPostMigrationXML.value

    $htmlCreateRoutingContact = get-date

    $telemetryFunctionStartTime = get-universalDateTime

    [int]$loopCounter = 0
    [boolean]$stopLoop = $FALSE

    if ($customRoutingDomain -eq "")
    {
        out-logfile -string "Calling new-routing contact without custom routing domain."
        do {
            try {
                new-routingContact -originalDLConfiguration $originalDLConfiguration -office365DlConfiguration $office365DLConfigurationPostMigration -globalCatalogServer $corevariables.globalCatalogWithPort.value -adCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod -customRoutingDomain $mailOnMicrosoftComDomain
    
                $stopLoop = $TRUE
            }
            catch {
                if ($loopCounter -gt 4)
                {
                    out-logfile -string $_ -isError:$TRUE
                }
                else {
                    start-sleepProgress -sleepString "Unable to create routing contact - try again." -sleepSeconds 5
    
                    $loopCounter = $loopCounter +1
                }
            }
        } while ($stopLoop -eq $FALSE)
    }
    else
    {
        out-logfile -string "Calling new-routingContact with custom domain."
        do {
            try {
                new-routingContact -originalDLConfiguration $originalDLConfiguration -office365DlConfiguration $office365DLConfigurationPostMigration -globalCatalogServer $corevariables.globalCatalogWithPort.value -adCredential $activeDirectoryCredential -customRoutingDomain $mailOnMicrosoftComDomain -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod
    
                $stopLoop = $TRUE
            }
            catch {
                if ($loopCounter -gt 4)
                {
                    out-logfile -string $_ -isError:$TRUE
                }
                else {
                    start-sleepProgress -sleepString "Unable to create routing contact - try again." -sleepSeconds 5
    
                    $loopCounter = $loopCounter +1
                }
            }
        } while ($stopLoop -eq $FALSE)
    }

    $stopLoop = $FALSE
    [int]$loopCounter = 0

    do {
        $tempMailArray = $originalDLConfiguration.mail.split("@")

        foreach ($member in $tempMailArray)
        {
            out-logfile -string ("Temp Mail Address Member: "+$member)
        }

        $tempMailAddress = $tempMailArray[0]+"-MigratedByScript"

        out-logfile -string ("Temp routing contact address: "+$tempMailAddress)

        $tempMailAddress = $tempMailAddress+"@"+$tempMailArray[1]

        out-logfile -string ("Temp routing contact address: "+$tempMailAddress)

        try {
            $routingContactConfiguration = Get-ADObjectConfiguration -groupSMTPAddress $tempMailAddress -globalCatalogServer $corevariables.globalCatalogWithPort.value -parameterSet $dlPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 

            $stopLoop=$TRUE
        }
        catch 
        {
            if ($loopCounter -gt 5)
            {
                out-logfile -string "Unable to obtain routing contact information post creation."
                out-logfile -string $_ -isError:$TRUE
            }
            else 
            {
                start-sleepProgress -sleepString "Unable to obtain routing contact after creation - sleep try again." -sleepSeconds 10
                $loopCounter = $loopCounter + 1                
            }
        }
    } while ($stopLoop -eq $FALSE)

    out-xmlFile -itemToExport $routingContactConfiguration -itemNameTOExport $xmlFiles.routingContactXML.value

    add-routingContactToGroup -originalDLConfiguration $originalDLConfiguration -routingContact $routingContactConfiguration -globalCatalogServer $corevariables.globalCatalogWithPort.value -activeDirectoryCredential $activeDirectoryCredential -activeDirectoryAuthenticationMethod $activeDirectoryAuthenticationMethod

    $routingContactConfiguration = Get-ADObjectConfiguration -groupSMTPAddress $tempMailAddress -globalCatalogServer $corevariables.globalCatalogWithPort.value -parameterSet $dlPropertySet -errorAction STOP -adCredential $activeDirectoryCredential 

    out-xmlFile -itemToExport $routingContactConfiguration -itemNameTOExport $xmlFiles.routingContactUpdatedXML.value

    $telemetryfunctionEndTime = get-universalDateTime

    $telemetryCreateRoutingContact = get-elapsedTime -startTime $telemetryfunctionStartTime -endTime $telemetryFunctionEndTime

    out-logfile -string "Calling function to disconnect all powershell sessions."

    disable-allPowerShellSessions

    $htmlEndTime = get-date

    $telemetryEndTime = get-universalDateTime
    $telemetryElapsedSeconds = get-elapsedTime -startTime $telemetryStartTime -endTime $telemetryEndTime

    # build the properties and metrics #
    $telemetryEventProperties = @{
        DLConversionV3Command = $telemetryEventName
        DLConversionV3Version = $telemetryDLConversionV3Version
        ExchangeOnlineVersion = $telemetryExchangeOnlineVersion
        MSGraphAuthentication = $telemetryMSGraphAuthentication
        MSGraphUsers = $telemetryMSGraphUsers
        MSGraphGroups = $telemetryMSGraphGroups
        AzureADVersion = $telemetryAzureADVersion
        OSVersion = $telemetryOSVersion
        MigrationStartTimeUTC = $telemetryStartTime
        MigrationEndTimeUTC = $telemetryEndTime
        MigrationErrors = $telemetryError
    }

    if (($allowTelemetryCollection -eq $TRUE) -and ($allowDetailedTelemetryCollection -eq $FALSE))
    {
        $telemetryEventMetrics = @{
            MigrationElapsedSeconds = $telemetryElapsedSeconds
            TimeToNormalizeDNs = $telemetryNormalizeDN
            TimeToValidateCloudRecipients = $telemetryValidateCloudRecipients
            TimeToCollectOnPremDependency = $telemetryDependencyOnPrem
            TimeToCollectOffice365Dependency = $telemetryCollectOffice365Dependency
            TimePendingRemoveDLOffice365 = $telemetryTimeToRemoveDL
            TimeToCreateOffice365DLComplete = $telemetryCreateOffice365DL
            TimeToCreateOffice365DLFirstPass = $telemetryCreateOffice365DLFirstPass
            TimeToReplaceOnPremDependency = $telemetryReplaceOnPremDependency
            TimeToReplaceOffice365Dependency = $telemetryReplaceOffice365Dependency
        }
    }
    elseif (($allowTelemetryCollection -eq $TRUE) -and ($allowDetailedTelemetryCollection -eq $TRUE))
    {
        $telemetryEventMetrics = @{
            MigrationElapsedSeconds = $telemetryElapsedSeconds
            TimeToNormalizeDNs = $telemetryNormalizeDN
            TimeToValidateCloudRecipients = $telemetryValidateCloudRecipients
            TimeToCollectOnPremDependency = $telemetryDependencyOnPrem
            TimeToCollectOffice365Dependency = $telemetryCollectOffice365Dependency
            TimePendingRemoveDLOffice365 = $telemetryTimeToRemoveDL
            TimeToCreateOffice365DLComplete = $telemetryCreateOffice365DL
            TimeToReplaceOnPremDependency = $telemetryReplaceOnPremDependency
            TimeToReplaceOffice365Dependency = $telemetryReplaceOffice365Dependency
            NumberOfGroupMembers = $exchangeDLMembershipSMTP.count
            NumberofGroupRejectSenders = $exchangeRejectMessagesSMTP.count
            NumberofGroupAcceptSenders = $exchangeAcceptMessagesSMTP.count
            NumberofGroupManagedBy = $exchangeManagedBySMTP.count
            NumberofGroupModeratedBy = $exchangeModeratedBySMTP.count
            NumberofGroupBypassModerators = $exchangeBypassModerationSMTP.count
            NumberofGroupGrantSendOnBehalfTo = $exchangeGrantSendOnBehalfToSMTP.count
            NumberofGroupSendAsOnGroup = $exchangeSendAsSMTP.Count
            NumberofOnPremsiesMemberOf = $allGroupsMemberOf.Count
            NumberofOnPremisesRejectSenders = $allGroupsReject.Count
            NumberofOnPremisesAcceptSenders = $allGroupsAccept.Count
            NumberofOnPremisesBypassModeration = $allGroupsBypassModeration.Count
            NumberofOnPremisesMailboxForwarding = $allUsersForwardingAddress.Count
            NumberofOnPrmiesesGrantSendBehalfTo = $allGroupsGrantSendOnBehalfTo.Count
            NumberofOnPremisesManagedBy = $allGroupsManagedBy.Count
            NumberofOnPremisesFullMailboxAccess = $allObjectsFullMailboxAccess.Count
            NumberofOnPremsiesSendAs = $allObjectSendAsAccess.Count
            NumberofOnPremisesFolderPermissions = $allMailboxesFolderPermissions.Count
            NumberofOnPremisesCoManagers = $allGroupsCoManagedByBL.Count
            NumberofOffice365Members = $allOffice365MemberOf.Count
            NumberofOffice365AcceptSenders = $allOffice365Accept.Count
            NumberofOffice365RejectSenders = $allOffice365Reject.Count
            NumberofOffice365BypassModeration = $allOffice365BypassModeration.Count
            NumberofOffice365ManagedBy = $allOffice365ManagedBy.Count
            NumberofOffice365GrantSendOnBehalf = $allOffice365GrantSendOnBehalfTo.Count
            NumberofOffice365ForwardingMailboxes= $allOffice365ForwardingAddress.Count
            NumberofOffice365FullMailboxAccess = $allOffice365FullMailboxAccess.Count
            NumberofOffice365SendAs = $allOffice365SendAsAccess.Count
            NumberofOffice365SendAsAccessOnGroup = $allOffice365SendAsAccessOnGroup.Count
            NumberofOffice365MailboxFolderPermissions = $allOffice365MailboxFolderPermissions.Count
        }
    }
    else 
    {
        $telemetryEventMetrics = @{}
    }

    if ($allowTelemetryCollection -eq $TRUE)
    {
        out-logfile -string "Telemetry1"
        out-logfile -string $traceModuleName
        out-logfile -string "Telemetry2"
        out-logfile -string $telemetryEventName
        out-logfile -string "Telemetry3"
        out-logfile -string $telemetryEventMetrics
        out-logfile -string "Telemetry4"
        out-logfile -string $telemetryEventProperties
        send-TelemetryEvent -traceModuleName $traceModuleName -eventName $telemetryEventName -eventMetrics $telemetryEventMetrics -eventProperties $telemetryEventProperties
    }

    generate-HTMLFile

    start-sleep -s 5

    if ($telemetryError -eq $TRUE)
    {
        out-logfile -string "" -isError:$TRUE
    }

    Start-ArchiveFiles -isSuccess:$TRUE -logFolderPath $logFolderPath

}