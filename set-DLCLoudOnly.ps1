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

        $functionURIType = "OnPremisesSyncBehavior"

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "START set-DLCloudOnly"
        Out-LogFile -string "********************************************************************************"
    
        out-logfile -string ("Acting on group: "+$office365DLConfiguration.externalDirectoryObjectID)

        try {
            $functionURI = get-graphURI -msGraphURL $msGraphURL -externalDirectoryObjectID $office365DLConfiguration.externalDirectoryObjectID -uriType $functionURIType -errorAction STOP
            out-logfile -string $functionURI
        }
        catch {
            out-logfile -string "Unable to obtain the graph URI"
            out-logfile -string $_ -isError:$TRUE
        }


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