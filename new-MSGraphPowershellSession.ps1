<#
    .SYNOPSIS

    This function creates the powershell session to msGraph AD.

    .DESCRIPTION

    This function creates the powershell session to msGraph AD.

    .PARAMETER msGraphADCredential

    The credential utilized to connect to msGraph ad.

    .PARAMETER msGraphCertificateThumbprint

    The certificate thumbprint for the associated msGraph application.

    .PARAMETER msGraphTenantID

    The tenant ID associated with the msGraph application.

    .PARAMETER msGraphApplicationID

    The application ID for msGraph management.

    .PARAMETER msGraphEnvironmentName

    The msGraph environment for the connection to msGraph ad.

	.OUTPUTS

    Powershell session to use for exchange online commands.

    .EXAMPLE

    new-msGraphADPowershellSession -msGraphADCredential $CRED -msGraphEnvironmentName NAME

    #>
    Function New-MSGraphPowershellSession
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [string]$msGraphCertificateThumbprint,
            [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [string]$msGraphApplicationID,
            [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [string]$msGraphTenantID,
            [Parameter(Mandatory = $true)]
            [string]$msGraphEnvironmentName,
            [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [string]$msGraphClientSecret,
            [Parameter(Mandatory = $true)]
            [array]$msGraphScopesRequired
        )

        #Output all parameters bound or unbound and their associated values.

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        #Define variables that will be utilzed in the function.
     
        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN NEW-msGraphADPowershellSession"
        Out-LogFile -string "********************************************************************************"

        if (($msGraphCertificateThumbprint -eq "") -and ($msGraphClientSecret -eq ""))
        {
            out-logfile -string "Attempting graph interactive authentication."

            try {
                connect-mgGraph -Scopes $msGraphScopesRequired -Environment $msGraphEnvironmentName -TenantId $msGraphTenantID -ErrorAction Stop
            }
            catch {
                out-logfile -string "Unable to connect to Microsoft Graph using Interactive Authentication"
                out-logfile -string $_
            }
        }
        elseif (($msGraphCertificateThumbprint -ne "") -or ($msGraphClientSecret -ne ""))
        {
            out-logfile -string "Either certificate or secret authentication specified."

            if ($msGraphCertificateThumbprint -ne "")
            {
                out-logfile -string "Certificate authentication specified."

                try {
                    Connect-MgGraph -TenantId $msGraphTenantID -CertificateThumbprint $msGraphCertificateThumbprint -ClientId $msGraphApplicationID -Environment $msGraphEnvironmentName -ErrorAction Stop
                }
                catch {
                    out-logfile -string "Unable to connect to Microsoft Graph using Certificate Authentication"
                }

            }
            else 
            {
                out-logfile -string "Client Secret authentication specified."

                out-logfile -string "Convert the secret to secure string."

                try {
                    $securedPasswordPassword = convertTo-SecureString -string $msGraphClientSecret -AsPlainText -Force -ErrorAction Stop
                }
                catch {
                    out-logfile -string "Unable to convert provided secret to secure string."
                    out-logfile -string $_ -isError:$TRUE
                }

                $clientSecretCredential = new-object -typeName System.Management.Automation.PSCredential -argumentList $msGraphApplicationID,$securedPasswordPassword

                try {
                    Connect-MgGraph -tenantID $msGraphTenantID -environment $msGraphEnvironmentName -ClientSecretCredential $clientSecretCredential -errorAction Stop
                }
                catch {
                    out-logfile -string "Unable to connect to Microsoft Graph using Client Secret Authentication"
                    out-logfile -string $_ -isError:$TRUE
                }
            }
        }

        Out-LogFile -string "END NEW-msGraphADPOWERSHELL SESSION"
        Out-LogFile -string "********************************************************************************"
    }
