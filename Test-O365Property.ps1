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
        [string]$isNotOk = "Yes"

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN TEST-O365Property"
        Out-LogFile -string "********************************************************************************"

        out-logfile -string "Obtain recipient information from Office 365."

        if ($member.externalDirectoryObjectID -ne $null)
        {
            out-logfile -string "External directory object ID specified - test."
            out-logfile -string $member.externalDirectoryObjectID

            $functionDirectoryObjectID=$member.externalDirectoryObjectID.Split("_")

            out-logfile -string $functionDirectoryObjectID[1]

            if ($member.recipientOrUser -eq "Recipient")
            {
                out-logfile -string "Member is recipient - query recipient."

                try {
                    $functionRecipient = get-o365recipient -identity $functionDirectoryObjectID[1] -errorAction STOP
                }   
                catch {
                    out-logfile -string "Unable to locate user by external directory object id."
                    $isTestError="Yes"
                }
            }
            else
            {
                out-logfile -string "Member is user - query user"

                try {
                    $functionRecipient = get-o365User -identity $functionDirectoryObjectID[1] -errorAction STOP
                }   
                catch {
                    out-logfile -string "Unable to locate user by external directory object id."
                    $isTestError="Yes"
                }
            }
        }
        elseif ($member.primarySMTPAddressOrUPN -ne $null)
        {
            out-logfile -string "Primary smtp address or upn specified - test."

            if ($member.recipientOrUser -eq "Recipient")
            {
                out-logfile -string "Member is recipient - query recipient."

                try {
                    $functionRecipient = get-o365recipient -identity $member.primarySMTPAddressOrUPN -errorAction STOP
                }   
                catch {
                    out-logfile -string "Unable to locate user by external directory object id."
                    $isTestError="Yes"
                }
            }
            else
            {
                out-logfile -string "Member is user - query user"

                try {
                    $functionRecipient = get-o365User -identity $member.primarySMTPAddressOrUPN -errorAction STOP
                }   
                catch {
                    out-logfile -string "Unable to locate user by external directory object id."
                    $isTestError="Yes"
                }
            }
        }

        if ($isTestError -eq "No")
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
            if ($isTestError -eq $TRUE)
            {
                out-logfile -string "The member was not located in Office 365 attribute - test recipient for presence."

                try {
                    $functionTest = get-o365Recipient -member $member -errorAction STOP
                }
                catch {
                    out-logfile -string "Unable to test the recipient presence in Office 365."
                    out-logfile -string $_ -isError:$true
                }

                if ($functionTestError -eq $isNotOk)
                {
                    out-logfile -string "Recipient not located in Office 365."
                    $member.isErrorMessageRecipient = "OFFICE_365_DEPENDENCY_NOT_FOUND_EXCEPTION: A group dependency was not found in Office 365.  Please either ensure the dependency is present or remove the dependency from the group."
                }
            }
        }

        Out-LogFile -string "END TEST-O365Property"
        Out-LogFile -string "********************************************************************************"    

        return $isTestError
    }