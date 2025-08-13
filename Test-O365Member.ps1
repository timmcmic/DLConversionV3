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
        [string]$isNotOk = "Yes"


        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN Test-O365Member"
        Out-LogFile -string "********************************************************************************"

        if ($membership.count -gt 0)
        {
            out-logfile -string "Only perform test if the count of objects in Office 365 > 0"

            if ($member.externalDirectoryObjectID -ne $NULL)
            {
                out-logfile -string "External directory object ID specified - test."
                out-logfile -string $member.externalDirectoryObjectID

                $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

                out-logfile -string $functionDirectoryObjectID[1]

                if ($membership.externalDirectoryObjectID.contains($functionDirectoryObjectID[1]))
                {
                    out-logfile -string "Member was located by external directory object id."
                }
                else 
                {                    
                    $isTestError="Yes"
                }
            }
            elseif ($member.primarySMTPAddressOrUPN -ne $NULL)
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
            else
            {
                out-logfile -string "Normalization failed to determine a UPN or External Directory Object ID."
                $isTestError="Yes"
            }
        }
        else 
        {
            out-logfile -string "To invoke this test the on premises attribute value has membership."
            out-logfile -string "To get here the corresponding attribute does not hvae membership - this is an error."
            $isTestError="Yes"
        }

        if ($isTestError -eq $isNotOk)
        {
            out-logfile -string "The member was not located in Office 365 attribute - test recipient for presence."

            try {
                $functionTest = Test-O365Recipient -member $member -errorAction STOP
            }
            catch {
                out-logfile -string "Unable to test the recipient presence in Office 365."
                out-logfile -string $_ -isError:$true
            }

            if ($functionTest -eq $isNotOk)
            {
                out-logfile -string "Recipient not located in Office 365."
                $member.isErrorMessageRecipient = "OFFICE_365_DEPENDENCY_NOT_FOUND_EXCEPTION: A group dependency was not found in Office 365.  Please either ensure the dependency is present or remove the dependency from the group."
            }
            else 
            {
                $member.isErrorMessageRecipient = "N/A"
            }
        }
        else 
        {
            $member.isErrorMessageRecipient = "N/A"
        }

        Out-LogFile -string "END Test-O365Member"
        Out-LogFile -string "********************************************************************************"    

        return $isTestError
    }