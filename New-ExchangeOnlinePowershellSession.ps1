<#
    .SYNOPSIS

    This function creates the powershell session to Exchange Online.

    .DESCRIPTION

    This function uses the exchange management shell v2 to utilize modern authentication to connect to exchange online.

    .PARAMETER exchangeOnlineCertificateThumbprint

    The user specified thumbprint if using certificate authentication for exchange online.

    .PARAMETER exchangeOnlineCredential

    The user specified credential for exchange online.

    .PARAMETER exchangeOnlineOrganiationName

    The onmicrosoft.com organization name.

    .PARAMETER exchangeOnlineAppID 

    The appilcation ID created in Azure for exchange online management.

    .PARAMETER exchangeOnlineEnvironmentName

    The Exchange online environment name if a non-commercial tenant is required.

	.OUTPUTS

    Powershell session to use for exchange online commands.

    .EXAMPLE

    New-ExchangeOnlinePowershellSession -exchangeOnlineCredentials $cred
    New-ExchangeOnlinePowershellSession -exchangeOnlineCertificate $thumbprint

    #>
    Function New-ExchangeOnlinePowershellSession
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [AllowNull()]
            [pscredential]$exchangeOnlineCredentials,
            [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [string]$exchangeOnlineCertificateThumbPrint,
            [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [string]$exchangeOnlineAppID,
            [Parameter(Mandatory = $true)]
            [AllowEmptyString()]
            [string]$exchangeOnlineOrganizationName,
            [Parameter(Mandatory = $true)]
            [string]$exchangeOnlineEnvironmentName,
            [Parameter(Mandatory = $true)]
            [string]$debugLogPath,
            [Parameter(Mandatory = $false)]
            [boolean]$isAudit=$FALSE
        )

        #Output all parameters bound or unbound and their associated values.

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        #Define variables that will be utilzed in the function.

        [string]$exchangeOnlineCommandPrefix="O365"
        
        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN NEW-EXCHANGEONLINEPOWERSHELLSESSION"
        Out-LogFile -string "********************************************************************************"

        #Log the parameters and variables for the function.

        if (($exchangeOnlineCredentials -eq $NULL) -and ($exchangeOnlineCertificateThumbPrint -eq "") -and ($exchangeOnlineAppID -eq "") -and ($exchangeOnlineOrganizationName -eq ""))
        {
            out-logfile -string "Attempting interactive authentication."

            try {
                Connect-ExchangeOnline -prefix $exchangeOnlineCommandPrefix -exchangeEnvironmentName $exchangeOnlineEnvironmentName -EnableErrorReporting -LogDirectoryPath $debugLogPath -LogLevel All -errorAction Stop
            }
            catch {
                out-logfile -string "Unable to complete Exchange Online Interactive Authentication"
                out-logfile -string $_ -isError:$TRUE
            }

        }
        elseif ($exchangeOnlineCredentials -ne ([System.Management.Automation.PSCredential]::Empty))
        {
            out-logfile -string "Attempting interactive authentication with credentials."

            try {
                Connect-ExchangeOnline -credential $exchangeOnlineCredentials -prefix $exchangeOnlineCommandPrefix -exchangeEnvironmentName $exchangeOnlineEnvironmentName -EnableErrorReporting -LogDirectoryPath $debugLogPath -LogLevel All -errorAction Stop
            }
            catch {
                out-logfile -string "Unable to complete Exchange Online Credential Authentication"
                out-logfile -string $_ -isError:$TRUE
            }
        }
        else 
        {
            out-logfile -string "Attempting certificate authentication."
            
            try {
                connect-exchangeOnline -certificateThumbPrint $exchangeOnlineCertificateThumbPrint -appID $exchangeOnlineAppID -Organization $exchangeOnlineOrganizationName -exchangeEnvironmentName $exchangeOnlineEnvironmentName -prefix $exchangeOnlineCommandPrefix -EnableErrorReporting -LogDirectoryPath $debugLogPath -LogLevel All
            }
            catch {
                out-logfile -string "Unable to complete Exchange Online Certificate Authentication"
                out-logfile -string $_ -isError:$TRUE
            }
        }
               
        Out-LogFile -string "The exchange online powershell session was created successfully."

        Out-LogFile -string "END NEW-EXCHANGEONLINEPOWERSHELLSESSION"
        Out-LogFile -string "********************************************************************************"
    }