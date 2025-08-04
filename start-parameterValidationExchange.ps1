<#
    .SYNOPSIS

    This function validates the parameters within the script.  Paramter validation is shared across functions.
    
    .DESCRIPTION

    This function validates the parameters within the script.  Paramter validation is shared across functions.

    #>
    Function start-parameterExchange
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

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "START start-parameterValidation"
        Out-LogFile -string "********************************************************************************"
        
        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END start-parameterValidation"
        Out-LogFile -string "********************************************************************************"
    }