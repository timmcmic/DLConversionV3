<#
    .SYNOPSIS

    This function obtains the DL membership of the Office 365 distribution group.

    .DESCRIPTION

    This function obtains the DL membership of the Office 365 distribution group.

    .PARAMETER GroupObjectID

    The Object ID to obtain membership values from Azure.

    .OUTPUTS

    Returns the membership array of the DL in Office 365.

    .EXAMPLE

    get-o365dlMembership -groupSMTPAddress Address

    #>
    Function Get-msGraphMembership
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            [string]$groupObjectID,
            [Parameter(Mandatory = $true)]
            $msGraphURL
        )

        #Output all parameters bound or unbound and their associated values.

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        #Declare function variables.

        $functionDLMembership=$NULL #Holds the return information for the group query.
        $functionURIType = "Members"

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN GET-msGraphMembership"
        Out-LogFile -string "********************************************************************************"

        try {
            $functionURI = get-GraphURI -msGraphURL $msGraphURL -externalDirectoryObjectID $groupObjectID -uriType $functionURIType -errorAction STOP
            out-logfile -string $functionURI
        }
        catch {
            out-logfile -string "Unable to obtain the graph URI."
            out-logfile -string $_ -isError:$TRUE
        }


        #Get the recipient using the exchange online powershell session.

        out-logfile -string "Attempting to obtain the MSgraph Group membership."

        try {
            $functionDLMembership = get-mgGroupMember -groupID $groupObjectID -all -errorAction STOP
            #$functionDLMembership = Invoke-MgGraphRequest -Method Get -Uri $functionURI -errorAction Stop
        }
        catch {
            out-logfile -string "Unable to obtain the azure group membership."
            out-logfile -string $_ -isError:$TRUE
        }

        if ($functionDLMembership.count -gt 0)
        {
            out-logfile -string $functionDLMembership
        }
        else
        {
            $functionDLMembers = @()
            out-logfile -string "No Azure AD Group members in the specified group."
        }
        
        Out-LogFile -string "END GET-MSGRAPHMEMBERSHIP"
        Out-LogFile -string "********************************************************************************"
        
        #Return the membership to the caller.
        
        return $functionDLMembership
    }