<#
    .SYNOPSIS

    This function tests if the recipient is found in Office 365.

    .DESCRIPTION

    This function tests to ensure a recipient is found in Office 365.

    .PARAMETER member

    The member to test for.

    .OUTPUTS

    None

    .EXAMPLE

    test-o365Recipient -member $member

    #>
    Function Test-O365Property
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $member
        )

        out-logfile -string "Output bound parameters..."

        #Output all parameters bound or unbound and their associated values.

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        #Declare local variables.

        [string]$isTestError="No"

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN TEST-O365Property"
        Out-LogFile -string "********************************************************************************"

        

        Out-LogFile -string "END TEST-O365RECIPIENT"
        Out-LogFile -string "********************************************************************************"    

        return $isTestError
    }