<#
    .SYNOPSIS

    This function tests if the recipient is found in Office 365.

    .DESCRIPTION

    This function tests to ensure a recipient is found in Office 365.

    .PARAMETER member

    The member to test for.

    .OUTPUTS

    None

    .EXAMPLE

    test-o365Recipient -member $member

    #>
    Function Test-O365Member
     {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $member,
            [Parameter(Mandatory = $true)]
            $memberShip
        )

        out-logfile -string "Output bound parameters..."

        #Output all parameters bound or unbound and their associated values.

        write-functionParameters -keyArray $MyInvocation.MyCommand.Parameters.Keys -parameterArray $PSBoundParameters -variableArray (Get-Variable -Scope Local -ErrorAction Ignore)

        #Declare local variables.

        [string]$isTestError="No"

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN Test-O365Member"
        Out-LogFile -string "********************************************************************************"

        if ($member.externalDirectoryObjectID -ne "")
        {
            out-logfile -string "External directory object ID specified - test."
            out-logfile -string $member.externalDirectoryObjectID

            if ($membership.externalDirectoryObjectID.contains($member.externalDirectoryObjectID))
            {
                out-logfile -string "Member was located by external directory object id."
            }
            else 
            {
                $isTestError="Yes"
            }
        }
        elseif ($member.primarySMTPAddressOrUPN -ne "")
        {
            out-logfile -string "Primary smtp address or upn specified - test."

            if ($membership.primarySMTPAddress.contains($member.primarySMTPAddressOrUPN))
            {
                out-logfile -string "Member was lcoated by primary smtp address or UPN."
            }
            else 
            {
                $isTestError="Yes"
            }
        }

        Out-LogFile -string "END Test-O365Member"
        Out-LogFile -string "********************************************************************************"    

        return $isTestError
    }