<#
    .SYNOPSIS

    This function gathers the group information from Azure Active Directory.

    .DESCRIPTION

    This function gathers the group information from Azure Active Directory.

    .PARAMETER office365DLConfiguration

    The Office 365 DL configuration for the group.

    .OUTPUTS

    Returns the information from the associated group from Azure AD>

    .EXAMPLE

    get-AzureADDLConfiguration -office365DLConfiguration $configuration

    #>
    Function get-msGraphDLConfiguration
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $office365DLConfiguration,
            [Parameter(Mandatory = $true)]
            $msGraphURL
        )

        #Output all parameters bound or unbound and their associated values.

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN get-msGraphDLConfiguration"
        Out-LogFile -string "********************************************************************************"

        #Get the recipient using the exchange online powershell session.

        $functionURI = $msGraphURL + "groups/"
        out-logfile -string $functionURI
        $functionURI = $functionURI + $office365DLConfiguration.externalDirectoryObjectID.tostring()
        out-logfile -string $functionURI
        
        try{
            #$functionDLConfiguration = get-mgGroup -groupID $office365DLConfiguration.externalDirectoryObjectID -errorAction STOP
            $functionDLConfiguration = Invoke-MgGraphRequest -Method Get -Uri $functionURI -errorAction STOP -debug 
        }
        catch {
            out-logfile -string "Unable to obtain group configuration from Azure Active Directory"
            out-logfile -string $_ -isError:$TRUE
        }

        Out-LogFile -string "END get-msGraphDLConfiguration"
        Out-LogFile -string "********************************************************************************"
        
        #This function is designed to open local and remote powershell sessions.
        #If the session requires import - for example exchange - return the session for later work.
        #If not no return is required.
        
        return $functionDLConfiguration
    }