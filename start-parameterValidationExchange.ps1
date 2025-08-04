<#
    .SYNOPSIS

    This function validates the parameters within the script.  Paramter validation is shared across functions.
    
    .DESCRIPTION

    This function validates the parameters within the script.  Paramter validation is shared across functions.

    #>
    Function start-parameterValidationExchange
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $exchangeOnlineCredential,
            [Parameter(Mandatory = $true)]
            $exchangeOnlineAppID,
            [Parameter(Mandatory = $true)]
            $exchangeOnlineCertificateThumbPrint,
            [Parameter(Mandatory = $true)]
            $exchangeOnlineOrganizationName
        ) 

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "START start-parameterValidationExchange"
        Out-LogFile -string "********************************************************************************"
        
        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END start-parameterValidationExchange"
        Out-LogFile -string "********************************************************************************"
    }