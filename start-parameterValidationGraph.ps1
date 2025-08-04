<#
    .SYNOPSIS

    This function validates the parameters within the script.  Paramter validation is shared across functions.
    
    .DESCRIPTION

    This function validates the parameters within the script.  Paramter validation is shared across functions.

    #>
    Function start-parameterValidationGraph
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $msGraphTenantID,
            [Parameter(Mandatory = $true)]
            $msGraphApplicationID,
            [Parameter(Mandatory = $true)]
            $msGraphCertificateThumbprint,
            [Parameter(Mandatory = $true)]
            $msGraphClientSecret
        )

        #Output all parameters bound or unbound and their associated values.

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN start-parameterValidationGraph"
        Out-LogFile -string "********************************************************************************"
    }