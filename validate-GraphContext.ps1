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

        foreach ($object in $functionGraphContext.psObject.properties)
        {
            if ($object.Value.count -gt 1)
            {
                foreach ($value in $object.Value)
                {
                    $string = ($object.name + " " + $value.tostring())
                    new-htmlListItem -text $string -fontSize 14
                }
            }
            elseif ($object.value -ne $NULL)
            {
                $string = ($object.name + " " + $object.value.tostring())
                new-htmlListItem -text $string -fontSize 14                            }
            else
            {
                $string = ($object.name)
                new-htmlListItem -text $string -fontSize 14
            }
        }

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END validate-GraphContext"
        Out-LogFile -string "********************************************************************************"
    }