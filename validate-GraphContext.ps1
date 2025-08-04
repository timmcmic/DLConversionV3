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

        try {
            $functionGraphContext = Get-MgContext -ErrorAction STOP
        }
        catch {
            out-logfile -string "Unable to obtain the graph context."
            out-logfile -string $_ -isError:$true
        }

        try {
            write-objectProperties -objectToWrite $functionGraphContext -errorAction STOP
        }
        catch {
            out-logfile -string "Unable to output graph context - this suggests previous unhandled exception connecting to graph." -isError:$TRUE
        }

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