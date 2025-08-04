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
            [AllowNull()]
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

        out-logfile -string "Validate that only a single Exchange credetial type is in use."

        if ($exchangeOnlineCredential -eq $NULL -and (($exchangeOnlineCertificateThumbPrint -eq "") -and ($exchangeOnlineAppID -eq "") -and ($exchangeOnlineOrganizationName -eq "")))
        {
            out-logfile -string "No Exchange Online credentials were specified - specify either credentials or application authentication." -isError:$TRUE
        }
        elseif ($exchangeOnlineCredential -ne $NULL -and (($exchangeOnlineCertificateThumbPrint -ne "") -or ($exchangeOnlineAppID -ne "") -or ($exchangeOnlineOrganizationName -ne "")))
        {
            out-logfile -string "Both an Exchange Online Credential and portions of Exchange Online Certificate Authenciation specified - choose one." -isError:$TRUE
        }
        else 
        {
            out-logfile -string "Only a single exchange online authentication method is specified."
        }

        out-logfile -string "If any portion of certificate authentication is specified and you've reached this point - test and advise."

        if (($exchangeOnlineAppID -eq "") -and (($exchangeOnlineOrganizationName -ne "") -and ($exchangeOnlineCertificateThumbPrint -ne "")))
        {
            out-logfile -string "Exchange Organization Name and Exchange Certificate Thumbprint required when specifing Exchange App ID."
        }


        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END start-parameterValidationExchange"
        Out-LogFile -string "********************************************************************************"
    }