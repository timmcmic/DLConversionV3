<#
    .SYNOPSIS

    This function outputs all of the items contained in an object and their associated values.

    .DESCRIPTION

    This function outputs all of the items contained in an object and their associated values.

    #>
    Function write-objectProperties
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $objectToWrite
        )

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
                out-logfile -string $string                        
            }
        }

        Out-LogFile -string "********************************************************************************"
    }