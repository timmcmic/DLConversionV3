<#
    .SYNOPSIS

    This function validates the graph context and ensures necessary scopes exist.

    .DESCRIPTION

    This function validates the graph context and ensures necessary scopes exist.

    #>
    Function validate-graphContext
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $msGraphScopesRequired
        )

        $functionGraphContext = $null

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "START validate-GraphContext"
        Out-LogFile -string "********************************************************************************"
    
        out-logfile -string "Obtain the graph context."

        $functionGraphContext = Get-MgContext

        out-logfile -string $functionGraphContext

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END validate-GraphContext"
        Out-LogFile -string "********************************************************************************"
    }