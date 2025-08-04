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

        if (($exchangeOnlineCredential -eq $NULL) -and (($exchangeOnlineCertificateThumbPrint -eq "") -and ($exchangeOnlineAppID -eq "") -and ($exchangeOnlineOrganizationName -eq "")))
        {
            out-logfile -string "No Exchange Online credentials specified - assume interactive authentication."
        }
        elseif (($exchangeOnlineCredential -ne $NULL) -and (($exchangeOnlineCertificateThumbPrint -ne "") -or ($exchangeOnlineAppID -ne "") -or ($exchangeOnlineOrganizationName -ne "")))
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
            out-logfile -string "Exchange Organization Name and Exchange Certificate Thumbprint required when specifing Exchange App ID." -isError:$TRUE
        }
        elseif (($exchangeOnlineCertificateThumbPrint -eq "") -and (($exchangeOnlineOrganizationName -ne "") -and ($exchangeOnlineAppID -ne "")))
        {
            out-logfile -string "Exchange Organization Name and Exchange App ID required when specifing Exchange Certificate Thumbprint." -isError:$TRUE
        }
        elseif (($exchangeOnlineOrganizationName -eq "") -and (($exchangeOnlineCertificateThumbPrint -ne "") -and ($exchangeOnlineAppID -ne "")))
        {
            out-logfile -string "Exchange Certificate Thumbprint and Exchange App ID required when specifing Exchange Organization Name." -isError:$TRUE
        }
        elseif (($exchangeOnlineAppID -ne "") -and (($exchangeOnlineOrganizationName -eq "") -and ($exchangeOnlineCertificateThumbPrint -eq "")))
        {
            out-logfile -string "Exchange Online organization name and certificate thumbprint required when specifying app id." -isError:$TRUE
        }
        elseif (($exchangeOnlineCertificateThumbPrint -ne "") -and (($exchangeOnlineOrganizationName -eq "") -and ($exchangeOnlineAppID -eq "")))
        {
            out-logfile -string "Exchange Online organization name and app id required when specifying certificate thumbprint." -isError:$TRUE
        }
        elseif (($exchangeOnlineOrganizationName -ne "") -and (($exchangeOnlineCertificateThumbPrint -eq "") -and ($exchangeOnlineAppID -eq "")))
        {
            out-logfile -string "Exchange certificate thumbprint and app id required when specifying organization name." -isError:$TRUE
        }
        else 
        {
            out-logfile -string "All components for Exchange Online certificate authentication specified."
        }

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END start-parameterValidationExchange"
        Out-LogFile -string "********************************************************************************"
    }