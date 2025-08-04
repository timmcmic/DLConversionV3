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

        out-logfile -string "Validate that only a single graph credetial type is in use."

        if (($msGraphClientSecret -ne "") -and ($msGraphCertificateThumbprint -ne ""))
        {
            out-logfile -string "A client secret and certificate thumbprint are specified - utilize only one method for graph app authentication." -isError:$TRUE
        }
        else 
        {
            out-logfile -string "Proceed with authentication method verification."
        }

        if ($msGraphClientSecret -ne "")
        {
            out-logfile -string "Client secret authentication specified."

            if (($msGraphTenantID -eq "") -and ($msGraphApplicationID -eq ""))
            {
                out-logfile -string "When specifying a graph client secret tenant ID and application ID are required." -isError:$TRUE
            }
            elseif(($msGraphTenantID -ne "") -and ($msGraphApplicationID -eq ""))
            {
                out-logfile -string "When specifying a graph client secret and tenant ID an application ID is required." -isError:$TRUE
            }
            elseif(($msGraphTenantID -eq "") -and ($msGraphApplicationID -ne ""))
            {
                out-logfile -string "When specifying a graph client secret and application ID the tenant ID is required." -isError:$TRUE
            }
            else 
            {
                out-logfile -string "All components for graph client secret authentication are present."
            }
        }
        elseif ($msGraphCertificateThumbprint -ne "")
        {
            out-logfile -string "Certificate authentication specified."

            if (($msGraphTenantID -eq "") -and ($msGraphApplicationID -eq ""))
            {
                out-logfile -string "When specifying a graph certificate thumbprint tenant ID and application ID are required." -isError:$TRUE
            }
            elseif(($msGraphTenantID -ne "") -and ($msGraphApplicationID -eq ""))
            {
                out-logfile -string "When specifying a graph certificate thumbprint and tenant ID an application ID is required." -isError:$TRUE
            }
            elseif(($msGraphTenantID -eq "") -and ($msGraphApplicationID -ne ""))
            {
                out-logfile -string "When specifying a graph certificate thumbprint and application ID the tenant ID is required." -isError:$TRUE
            }
            else 
            {
                out-logfile -string "All components for graph certificate authentication are present."
            }
        }

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN start-parameterValidationGraph"
        Out-LogFile -string "********************************************************************************"
    }