
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
        [boolean]$testRecipientHealth=$true
    )

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

    $telemetryInfo = @{
        telemetryDLConversionV3Version = $NULL
        telemetryExchangeOnlineVersion = $NULL
        telemetryAzureADVersion = $NULL
        telemetryMSGraphAuthentication = $NULL
        telemetryMSGraphUsers = $NULL
        telemetryMSGraphGroups = $NULL
        telemetryActiveDirectoryVersion = $NULL
        telemetryOSVersion = (Get-CimInstance Win32_OperatingSystem).version
        telemetryStartTime = get-universalDateTime
        telemetryEndTime = $NULL
        telemetryElapsedSeconds = [double]0
        telemetryEventName = "Start-DistributionListMigration-V3"
        telemetryFunctionStartTime = $NULL
        telemetryFunctionEndTime = $NULL
        telemetryNormalizeDN = [double]0
        telemetryValidateCloudRecipients = [double]0
        telemetryDependencyOnPrem = [double]0
        telemetryCollectOffice365Dependency = [double]0
        telemetryTimeToRemoveDL = [double]0
        telemetryCreateOffice365DL = [double]0
        telemetryCreateOffice365DLFirstPass = [double]0
        telemetryReplaceOnPremDependency = [double]0
        telemetryReplaceOffice365Dependency = [double]0
        telemetryError = [boolean]$FALSE
    }

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
        msGraphDLMembershipXML = @{"Value" = "msGraphADDLMembership" ; "Description" = "Export XML file holding the membership of the Azure AD group"}
        preCreateErrorsXML = @{"value" = "preCreateErrors" ; "Description" = "Export XML of all precreate errors for group to be migrated."}
        testOffice365ErrorsXML = @{"value" = "testOffice365Errors" ; "Description" = "Export XML of all tested recipient errors in Offic3 365."}
        office365DLMembership = @{"Value" = "office365DLMembership" ; "Description" = "Original Office 365 DL Membership"}
    }

    #On premises variables for the distribution list to be migrated.

    $originalDLConfiguration=$NULL #This holds the on premises DL configuration for the group to be migrated.
    $originalAzureADConfiguration=$NULL #This holds the azure ad DL configuration
    $originalDLConfigurationUpdated=$NULL #This holds the on premises DL configuration post the rename operations.
    $routingContactConfig=$NULL #Holds the mail routing contact configuration.
    $routingDynamicGroupConfig=$NULL #Holds the dynamic distribution list configuration used for mail routing.
    $routingContactConfiguration=$NULL #This is the empty routing contact configuration.
    [array]$exchangeDLMembershipSMTP=@() #Array of DL membership from AD.
    [array]$exchangeRejectMessagesSMTP=@() #Array of members with reject permissions from AD.
    [array]$exchangeAcceptMessagesSMTP=@() #Array of members with accept permissions from AD.
    [array]$exchangeManagedBySMTP=@() #Array of members with manage by rights from AD.
    [array]$exchangeModeratedBySMTP=@() #Array of members  with moderation rights.
    [array]$exchangeBypassModerationSMTP=@() #Array of objects with bypass moderation rights from AD.
    [array]$exchangeGrantSendOnBehalfToSMTP=@() #Array of objects with grant send on behalf to normalized SMTP
    [array]$exchangeSendAsSMTP=@() #Array of objects wtih send as rights normalized SMTP

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
    $office365DLMembership=$NULL
    $office365DLMembershipPostMigration=$NULL #This holds the Office 365 DL membership information post migration
    $msGraphURL = ""

    $dlPropertySet = '*' #Clear all properties of a given object

    [array]$global:preCreateErrors=@()
    [array]$global:testOffice365Errors=@()
    [array]$global:generalErrors=@()
    [string]$isTestError="No"

    #Initilize the log file.

    $global:logFile=$NULL #This is the global variable for the calculated log file name
    [string]$global:staticFolderName="\DLMigration\"
    new-LogFile -groupSMTPAddress $groupSMTPAddress.trim() -logFolderPath $logFolderPath
    $traceFilePath = $logFolderPath + $global:staticFolderName

    out-logfile -string ("Log File: "+$global:logFile)
    out-logfile -string ("Trace File: "+$traceFilePath)

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

    out-logfile -string ("Runtime start UTC: " + $telemetryInfo.telemetryStartTime.ToString())

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

    $telemetryInfo.telemetryExchangeOnlineVersion = Test-PowershellModule -powershellModuleName $corevariables.exchangeOnlinePowershellModuleName.value -powershellVersionTest:$TRUE

    Out-LogFile -string "Calling Test-PowerShellModule to validate the Active Directory is installed."

    $telemetryInfo.telemetryActiveDirectoryVersion = Test-PowershellModule -powershellModuleName $corevariables.activeDirectoryPowershellModuleName.value

    out-logfile -string "Calling Test-PowershellModule to validate the DL Conversion Module version installed."

    $telemetryInfo.telemetryDLConversionV3Version = Test-PowershellModule -powershellModuleName $corevariables.dlConversionPowershellModule.value -powershellVersionTest:$TRUE

    out-logfile -string "Calling Test-PowershellModule to validate the Microsoft Graph Authentication versions installed."

    $telemetryInfo.telemetryMSGraphAuthentication = test-powershellModule -powershellmodulename $corevariables.msgraphauthenticationpowershellmodulename.value -powershellVersionTest:$TRUE

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

    $msGraphDLMembership = get-msGraphMembership -groupobjectID $msGraphDLConfiguration.id -msGraphURL $msGraphURL -errorAction STOP

    out-xmlFile -itemToExport $msGraphDLMembership -itemNameToExport $xmlFiles.msGraphDLMembershipXML.value

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
    $telemetryInfo.FunctionStartTime = get-universalDateTime

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
                    $exchangeDLMembershipSMTP+=$normalizedTest
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

    $telemetryInfo.FunctionEndTime = get-universalDateTime

    $telemetryNormalizeDN = get-elapsedTime -startTime $telemetryInfo.FunctionStartTime -endTime $telemetryInfo.FunctionEndTime

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

    $htmlStartCloudValidation = get-date

    $telemetryInfo.FunctionStartTime = get-universalDateTime

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

                        $global:testOffice365Errors += $member
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

                        $global:testOffice365Errors += $member
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

                        $global:testOffice365Errors += $member
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

                        $global:testOffice365Errors += $member
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

                        $global:testOffice365Errors += $member
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

                        $global:testOffice365Errors += $member
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

                        $global:testOffice365Errors += $member
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

        out-logfile -string "Begin evaluation all members with send as rights."
    }

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "END VALIDATE RECIPIENTS IN CLOUD"
    Out-LogFile -string "********************************************************************************"

    $telemetryInfo.FunctionEndTime = get-universalDateTime

    $telemetryValidateCloudRecipients = get-elapsedTime -startTime $telemetryFunctionStartTime -endTime $telemetryFunctionEndTime

    out-logfile -string ("Time to validate recipients in cloud: "+ $telemetryValidateCloudRecipients.toString())
}