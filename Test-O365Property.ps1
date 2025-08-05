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
    Function Test-O365Property
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
        $functionRecipient=$NULL

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN TEST-O365Property"
        Out-LogFile -string "********************************************************************************"

        out-logfile -string "Obtain recipient information from Office 365."

        if (($member.externalDirectoryObjectID -ne $null) -and ($member.isMigrated -eq $FALSE) -and ($member.RecipientType -ne "msExchDynamicDistributionList"))
        {
            out-logfile -string "External directory object ID specified - test."
            out-logfile -string $member.externalDirectoryObjectID

            $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

            out-logfile -string $functionDirectoryObjectID[1]

            try {
                $functionRecipient = get-o365recipient -identity $functionDirectoryObjectID[1] -errorAction STOP
            }
            catch {
                out-logfile -string "Unable to locate user by external directory object id."
                $isTestError="Yes"
            }
        }
        elseif (($member.primarySMTPAddressOrUPN -ne $null) -and ($member.isMigrated -eq $FALSE) -and ($member.RecipientType -ne "msExchDynamicDistributionList")) 
        {
            out-logfile -string "Primary smtp address or upn specified - test."

            try {
                $functionRecipient = get-o365recipient -identity $member.primarySMTPAddressOrUPN -errorAction STOP
            }
            catch {
                out-logfile -string "Unable to locate user by primary SMTP address or UPN."
                $isTestError="Yes"
            }
        }

        if (($isTestError -eq "No") -and ($member.isMigrated -eq $FALSE) -and ($member.RecipientType -ne "msExchDynamicDistributionList"))
        {
            out-logfile -string "Previous errors not encountered - test further."
            out-logfile -string $functionRecipient.name

            if ($membership.count -gt 0)
            {
                 if ($membership.contains($functionRecipient.name))
                {
                    out-logfile -string "User was located successfully."
                }
                else 
                {
                    out-logfile -string "User was not located successfully."
                    $isTestError="Yes"
                }
            }   
        }
        else 
        {
            Out-logfile -string "No further testing required."
        }

        Out-LogFile -string "END TEST-O365Property"
        Out-LogFile -string "********************************************************************************"    

        return $isTestError
    }