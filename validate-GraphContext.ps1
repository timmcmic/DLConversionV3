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

        write-objectProperties -objectToWrite $functionGraphContext

        out-logfile -string "Testing graph scopes."

        foreach ($scope in $msGraphScopesRequired)
        {
            out-logfile -string ("Testing scope: "+$scope)

            if ($functionGraphContext.Scopes.contains($scope))
            {
                out-logfile -string "Required scope located - proceed."
            }
            else 
            {
                Out-logfile -string ("Graph Scope Required and Missing: "+$scope) -isError:$TRUE
            }
        }

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END validate-GraphContext"
        Out-LogFile -string "********************************************************************************"
    }