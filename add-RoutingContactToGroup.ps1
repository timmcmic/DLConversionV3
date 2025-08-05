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
        $originalDLConfiguration
    )

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "Start add-RoutingContactToGroup"
    Out-LogFile -string "********************************************************************************"


    foreach ($key in $hashtable.GetEnumerator())
    {
        out-logfile -string ("Key: "+$key.name+" is "+$key.Value.Description+" with value "+$key.Value.Value)
    }      

    out-LogFile -string "********************************************************************************"
    Out-LogFile -string "End add-RoutingContactToGroup"
    Out-LogFile -string "********************************************************************************"
}