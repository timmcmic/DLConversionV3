<#
    .SYNOPSIS

    This function validates the graph context and ensures necessary scopes exist.

    .DESCRIPTION

    This function validates the graph context and ensures necessary scopes exist.

    #>
    Function validate-ADRecycleBinEnabled
    {
        $functionTest = $null

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

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END validate-ADRecycleBinDisabled"
        Out-LogFile -string "********************************************************************************"

        return $functionTest
    }