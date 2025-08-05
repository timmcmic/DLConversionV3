<#
    .SYNOPSIS

    This function outputs all of the items contained in an object and their associated values.

    .DESCRIPTION

    This function outputs all of the items contained in an object and their associated values.

    #>
    Function set-DLCloudOnly
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $msGraphURL,
            [Parameter(Mandatory = $true)]
            $office365DLConfiguration
        )

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "START set-DLCloudOnly"
        Out-LogFile -string "********************************************************************************"
    
        out-logfile -string ("Acting on group: "+$office365DLConfiguration.externalDirectoryObjectID)

        $functionURI = $msGraphURL + "groups/"
        out-logfile -string $functionURI
        $functionURI = $functionURI + $office365DLConfiguration.externalDirectoryObjectID
        out-logfile -string $functionURI
        $functionURI = $functionURI + "/onPremisesSyncBehavior".
        out-logfile -string $functionURI

        try {
            Invoke-MgGraphRequest -Method Patch -Uri $functionURI -body @{isCloudManaged=$true} -errorAction STOP
        }
        catch {
            out-logfile -string "Unable to set the group to cloud only."
            out-logfile -string $_
        }
        
        Out-LogFile -string "END set-DLCloudOnly"
        Out-LogFile -string "********************************************************************************"
    }