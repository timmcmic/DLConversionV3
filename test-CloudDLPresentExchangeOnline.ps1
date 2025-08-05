<#
    .SYNOPSIS

    This function loops until we detect that the cloud DL is no longer present.
    
    .DESCRIPTION

    This function loops until we detect that the cloud DL is no longer present.

    .PARAMETER groupSMTPAddress

    The SMTP Address of the group.

    .OUTPUTS

    None

    .EXAMPLE

    test-CloudDLPresent -groupSMTPAddress SMTPAddress

    #>
    Function test-CloudDLPresentExchangeOnline
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

        $functionRecipient = $NULL

        #Start function processing.

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN test-CloudDLPresentExchangeOnline"
        Out-LogFile -string "********************************************************************************"

        try {
            $functionRecipient = get-o365Recipient -identity $groupSMTPAddress -errorAction STOP
        }
        catch {
            out-logfile -string "Unable to obtain the Exchange Online distribution list."
            out-logfile -string $_ -isError:$TRUE
        } 

        do 
        {
            start-sleepProgress -sleepString "Group still directory synchronized in Exchange Online - sleep for 30 seconds - try again." -sleepSeconds 30
            $functionRecipient = get-o365Recipient -identity $groupSMTPAddress -errorAction STOP

        } while ($functionRecipient.isDirSynced -eq $true)

        Out-LogFile -string "END test-CloudDLPresentExchangeOnline"
        Out-LogFile -string "********************************************************************************"
    }