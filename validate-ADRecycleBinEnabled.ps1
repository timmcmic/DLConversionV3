<#
    .SYNOPSIS

    This function validates the graph context and ensures necessary scopes exist.

    .DESCRIPTION

    This function validates the graph context and ensures necessary scopes exist.

    #>
    Function validate-ADRecycleBinDisabled
    {
        $functionTest = $null
        $functionGood = "Good"
        $functionBad = "Bad"
        $functionResult = ""

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "START validate-ADRecycleBinDisabled"
        Out-LogFile -string "********************************************************************************"

        out-logfile -string "Evaluating enabled scopes."

        try {
            $functionTest = Get-ADOptionalFeature -Filter {Name -eq "Recycle Bin Feature"} -errorAction Stop
        }
        catch {
            out-logfile -string "Unable to obtain the ad recycle bin info."
            out-logfile -string $_ -isError:$TRUE
        }
    

        if (($functionTest.enabledScopes.count -gt 0) -or ($funcionTest.enabledScopes -ne $null))
        {
            out-logfile -string "Enabled scopes are present - this signifies recycle bin is enabled."

            foreach ($scope in $functionTest.enabledScopes)
            {
                out-logfile -string $scope
            }

            $functionResult = $functionGood
        }
        else 
        {
            out-logfile -string "Recycle bin is not enabled - not good for this scenario."

            $functionResult = $functionBad
        }

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END validate-ADRecycleBinDisabled"
        Out-LogFile -string "********************************************************************************"

        return $functionResult
    }