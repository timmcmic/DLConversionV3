<#
.SYNOPSIS

This function outputs all of the parameters from a function to the log file for review.

.DESCRIPTION

This function outputs all of the parameters from a function to the log file for review.

#>
Function get-graphURI
{
    [cmdletbinding()]

    Param
    (
        [Parameter(Mandatory = $true)]
        $msGraphURL,
        [Parameter(Mandatory = $true)]
        $externalDirectoryObjectID,
        [Parameter(Mandatory = $true)]
        $uriType
    )

    $functionOnPremSyncBehavior = "OnPremisesSyncBehavior"
    $functionConfiguration = "Configuration"
    $functionMembers = "Members"

    Out-LogFile -string "********************************************************************************"
    Out-LogFile -string "Start get-graphURI"
    Out-LogFile -string "********************************************************************************"

    if ($uriType -eq $functionOnPremisesSyncBehavior)
    {
        $functionURI = $msGraphURL + "groups/"
        out-logfile -string $functionURI
        $functionURI = $functionURI + $externalDirectoryObjectID
        out-logfile -string $functionURI
        $functionURI = $functionURI + "/onPremisesSyncBehavior"
        out-logfile -string $functionURI
    }

    if ($uriType -eq $functionConfiguration)
    {
        $functionURI = $msGraphURL + "groups/"
        out-logfile -string $functionURI
        $functionURI = $functionURI + $office365DLConfiguration.externalDirectoryObjectID.tostring()
        out-logfile -string $functionURI

    }

    if ($uriType -eq $functionMembers)
    {
        $functionURI = $msGraphURL + "groups/"
        out-logfile -string $functionURI
        $functionURI = $functionURI + $groupObjectID
        out-logfile -string $functionURI
        $functionURI = $functionURI + "/members"
        out-logfile -string $functionURI
    }

    out-LogFile -string "********************************************************************************"
    Out-LogFile -string "End get-graphURI"
    Out-LogFile -string "********************************************************************************"

    return $functionURI
}