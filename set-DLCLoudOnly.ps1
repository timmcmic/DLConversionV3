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
    
        
        
        Out-LogFile -string "END set-DLCloudOnly"
        Out-LogFile -string "********************************************************************************"
    }