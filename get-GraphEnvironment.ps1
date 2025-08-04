<#
    .SYNOPSIS

    This function returns the approrpiate graph environment

    .DESCRIPTION

    This function returns the approrpiate graph environment

    #>
    Function get-GraphEnvironment
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $msGraphEnvironmentName,
            [Parameter(Mandatory = $false)]
            $useBeta=$false
        )

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN get-GraphEnvironment"
        Out-LogFile -string "********************************************************************************"
    
        foreach ($object in $objectToWrite.psObject.properties)
        {
            if ($object.Value.count -gt 1)
            {
                foreach ($value in $object.Value)
                {
                    $string = ($object.name + " " + $value.tostring())
                    out-logfile -string $string
                }
            }
            elseif ($object.value -ne $NULL)
            {
                $string = ($object.name + " " + $object.value.tostring())
                out-logfile -string $string                        }
            else
            {
                $string = ($object.name)
                out-logfile -string $string
            }
        }

        Out-LogFile -string "END get-GraphEnvironment"
        Out-LogFile -string "********************************************************************************"
    }