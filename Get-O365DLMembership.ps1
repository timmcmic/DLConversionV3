<#
    .SYNOPSIS

    This function obtains the DL membership of the Office 365 distribution group.

    .DESCRIPTION

    This function obtains the DL membership of the Office 365 distribution group.

    .PARAMETER GroupSMTPAddress

    The mail attribute of the group to search.

    .OUTPUTS

    Returns the membership array of the DL in Office 365.

    .EXAMPLE

    get-o365dlMembership -groupSMTPAddress Address

    #>
    Function Get-o365DLMembership
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$groupSMTPAddress
        )

        #Output all parameters bound or unbound and their associated values.

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        #Declare function variables.

        $functionDLMembership=$NULL #Holds the return information for the group query.

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN GET-O365DLMEMBERSHIP"
        Out-LogFile -string "********************************************************************************"

        #Get the recipient using the exchange online powershell session.

        try {
            $functionDLMembership=@(get-O365DistributionGroupMember -identity $groupSMTPAddress -resultsize unlimited -errorAction STOP | select-object Identity,Name,Alias,ExternalDirectoryObjectId,EmailAddresses,PrimarySMTPAddress,ExternalEmailAddress,DisplayName,RecipientType,RecipientTypeDetails,ExchangeGuid)
        }
        catch {
            out-logfile -string "Unable to obtain Exchange Online DL membership."
            out-logfile -string $_ -isError:$TRUE
        }
        
        Out-LogFile -string "END GET-O365DLMEMBERSHIP"
        Out-LogFile -string "********************************************************************************"
        
        #Return the membership to the caller.
        
        return $functionDLMembership
    }