<#
.SYNOPSIS

This function adds the routing contact to the on premises group to maintain mail flow.

.DESCRIPTION

This function adds the routing contact to the on premises group to maintain mail flow.


#>
Function add-RoutingContactToGroup
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory = $true)]
        $routingContact,
        [Parameter(Mandatory = $true)]
        $originalDLConfiguration,
        [Parameter(Mandatory = $true)]
        $globalCatalogServer,
        [Parameter(Mandatory = $true)]
        $activeDirectoryCredential,
        [Parameter(Mandatory = $true)]
        $activeDirectoryAuthenticationMethod
    )

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "Start add-RoutingContactToGroup"
    Out-LogFile -string "********************************************************************************"

    out-logfile -string $originalDLConfiguration.distinguishedName
    $functionDistinguishName = $originalDLConfiguration.distinguishedName.tostring()
    out-logfile -string $routingContact.distinguishedName
    $functionDistinguishedNameContact = $routingContact.distinguishedName.tostring()

    try {
        add-adGroupMember -identity $functionDistinguishName -Members $functionDistinguishedNameContact -Credential $activeDirectoryCredential -Server $globalCatalogServer -authType $activeDirectoryAuthenticationMethod -errorAction STOP
        out-logfile -string "Routing contact successfully added as group member."
    }
    catch {
        out-logfile -string "Unable to add the routing contact as a member of the original on premises group."
        out-logfile -string $_ -isError:$true
    }

    out-LogFile -string "********************************************************************************"
    Out-LogFile -string "End add-RoutingContactToGroup"
    Out-LogFile -string "********************************************************************************"
}