<#
    .SYNOPSIS

    This function validates the parameters within the script.  Paramter validation is shared across functions.
    
    .DESCRIPTION

    This function validates the parameters within the script.  Paramter validation is shared across functions.

    #>
    Function start-parameterValidation
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true,ParameterSetName = 'ExchangeOnline')]
            [AllowNull()]
            $exchangeOnlineCredential,
            [Parameter(Mandatory = $true,ParameterSetName = 'ExchangeOnlineCertAuth')]
            [AllowNull()]
            $exchangeOnlineCertificateThumbprint,
            [Parameter(Mandatory = $true,ParameterSetName = 'ExchangeOnlineCertAuth')]
            [AllowNull()]
            $exchangeOnlineOrganizationName,
            [Parameter(Mandatory = $true,ParameterSetName = 'ExchangeOnlineCertAuth')]
            [AllowNull()]
            $exchangeOnlineAppID,
            [Parameter(Mandatory = $true,ParameterSetName = 'ActiveDirectory')]
            [AllowNull()]
            $activeDirectoryCredential,
            [Parameter(Mandatory = $true,ParameterSetName = 'msGraphCertAuth')]
            [AllowNull()]
            $msGraphCertificateThumbprint,
            [Parameter(Mandatory = $true,ParameterSetName = 'msGraphCertAuth')]
            [Parameter(Mandatory = $true,ParameterSetName = 'msGraphClientSecretAuth')]
            [AllowNull()]
            $msGraphTenantID,
            [Parameter(Mandatory = $true,ParameterSetName = 'msGraphCertAuth')]
            [Parameter(Mandatory = $true,ParameterSetName = 'msGraphClientSecretAuth')]
            [AllowNull()]
            $msGraphApplicationID,
            [Parameter(Mandatory = $true,ParameterSetName = 'msGraphCertAuth')]
            [Parameter(Mandatory = $true,ParameterSetName = 'msGraphClientSecretAuth')]
            [AllowNull()]
            $msGraphClientSecret
        )

        #Output all parameters bound or unbound and their associated values.

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        $functionParameterSetName = $PsCmdlet.ParameterSetName
        $exchangeOnlineParameterSetName = "ExchangeOnline"
        $exchangeOnlineParameterSetNameCertAuth = "ExchangeOnlineCertAuth"
        $msGraphParameterSetNameCertAuth = "MSGraphCertAuth"
        $msGraphParameterSetNameSecretAuth = "MsGraphClientSecretAuth"
        $activeDirectoryParameterSetName = "ActiveDirectory"

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN start-parameterValidation"
        Out-LogFile -string "********************************************************************************"

        out-logfile -string ("The parameter set name for validation: "+$functionParameterSetName)

        if ($functionParameterSetName -eq $activeDirectoryParameterSetName)
        {
            test-credentials -credentialsToTest $activeDirectoryCredential
        }

        if ($functionParameterSetName -eq $msGraphParameterSetNameCertAuth)
        {
            if (($msGraphCertificateThumbprint -ne "") -and ($msGraphTenantID -eq "") -and ($msGraphApplicationID -eq ""))
            {
                out-logfile -string "The msGraph tenant ID and msGraph App ID are required when using certificate authentication to msGraph." -isError:$TRUE
            }
            elseif (($msGraphCertificateThumbprint -ne "") -and ($msGraphTenantID -ne "") -and ($msGraphApplicationID -eq ""))
            {
                out-logfile -string "The msGraph app id is required to use certificate authentication to msGraph." -isError:$TRUE
            }
            elseif (($msGraphCertificateThumbprint -ne "") -and ($msGraphTenantID -eq "") -and ($msGraphApplicationID -ne ""))
            {
                out-logfile -string "The msGraph tenant ID is required to use certificate authentication to msGraph." -isError:$TRUE
            }
            elseif (($msGraphCertificateThumbprint -eq "") -and ($msGraphTenantID -eq "") -and ($msGraphApplicationID -ne ""))
            {
                out-logfile -string "No components of msGraph Cert Authentication were provided - this is not necessarily an issue."
            }
            else 
            {
                out-logfile -string "All components necessary for MSGraphCertificate Authentication provided."    
            }
        }

        if ($functionParameterSetName -eq $msGraphParameterSetNameSecretAuth)
        {
            if (($msGraphClientSecret -ne "") -and ($msGraphTenantID -eq "") -and ($msGraphApplicationID -eq ""))
            {
                out-logfile -string "The msGraph tenant ID and msGraph App ID are required when using certificate authentication to msGraph." -isError:$TRUE
            }
            elseif (($msGraphClientSecret -ne "") -and ($msGraphTenantID -ne "") -and ($msGraphApplicationID -eq ""))
            {
                out-logfile -string "The msGraph app id is required to use certificate authentication to msGraph." -isError:$TRUE
            }
            elseif (($msGraphClientSecret -ne "") -and ($msGraphTenantID -eq "") -and ($msGraphApplicationID -ne ""))
            {
                out-logfile -string "The msGraph tenant ID is required to use certificate authentication to msGraph." -isError:$TRUE
            }
            elseif (($msGraphClientSecret -eq "") -and ($msGraphTenantID -eq "") -and ($msGraphApplicationID -eq ""))
            {
                out-logfile -string "No components of msGraph Cert Authentication were provided - this is not necessarily an issue."
            }
            else 
            {
                out-logfile -string "All components necesary for MSGraphClientSecret Authentication provided."    
            }
        }

        if ($functionParameterSetName -eq $exchangeOnlineParameterSetNameCertAuth)
        {
            if (($exchangeOnlineCertificateThumbPrint -ne "") -and ($exchangeOnlineOrganizationName -eq "") -and ($exchangeOnlineAppID -eq ""))
            {
                out-logfile -string "The exchange organization name and application ID are required when using certificate thumbprint authentication to Exchange Online." -isError:$TRUE
            }
            elseif (($exchangeOnlineCertificateThumbPrint -ne "") -and ($exchangeOnlineOrganizationName -ne "") -and ($exchangeOnlineAppID -eq ""))
            {
                out-logfile -string "The exchange application ID is required when using certificate thumbprint authentication." -isError:$TRUE
            }
            elseif (($exchangeOnlineCertificateThumbPrint -ne "") -and ($exchangeOnlineOrganizationName -eq "") -and ($exchangeOnlineAppID -ne ""))
            {
                out-logfile -string "The exchange organization name is required when using certificate thumbprint authentication." -isError:$TRUE
            }
            elseif (($exchangeOnlineCertificateThumbPrint -eq "") -and ($exchangeOnlineOrganizationName -eq "") -and ($exchangeOnlineAppID -eq ""))
            {
                out-logfile -string "No components of certificate authentication were specified.  This is not necessary an error."
            }
            else 
            {
                out-logfile -string "All components necessary for Exchange certificate thumbprint authentication were specified."    
            }
        }

        if (($functionParameterSetName -eq $exchangeOnlineParameterSetName) -or ($functionParameterSetName -eq $exchangeOnlineParameterSetNameCertAuth))
        {
            if (($exchangeOnlineCredential -ne $NULL) -and ($exchangeOnlineCertificateThumbPrint -ne ""))
            {
                Out-LogFile -string "ERROR:  Only one method of cloud authentication can be specified.  Use either cloud credentials or cloud certificate thumbprint." -isError:$TRUE
            }
            else
            {
                Out-LogFile -string "Only one method of Exchange Online authentication specified."

                if ($functionParamterSetName -eq $exchangeOnlineParameterSetName)
                {
                    out-logfile -string "Validating the exchange online credential array"

                    test-credentials -credentialsToTest $exchangeOnlineCredential
                }
            } 
        }
        
        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "END start-parameterValidation"
        Out-LogFile -string "********************************************************************************"
    }