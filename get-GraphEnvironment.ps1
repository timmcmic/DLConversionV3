<#
    .SYNOPSIS

    This function returns the approrpiate graph environment

    .DESCRIPTION

    This function returns the approrpiate graph environment

    #>
    Function get-GraphEnvironment
    {
        [cmdletbinding()]

        Param
        (
            [Parameter(Mandatory = $true)]
            $msGraphEnvironmentName,
            [Parameter(Mandatory = $false)]
            $useBeta=$false
        )

        $msGraphURLGlobal = "https://graph.microsoft.com/v1.0/"
        $msGraphURLUSGov = "https://graph.microsoft.us/v1.0/"
        $msGraphURLUSDoD = "https://dod-graph.microsoft.us/v1.0/"
        $msGraphURLChina = "https://microsoftgraph.chinacloudapi.cn/v1.0/"
        $msGraphBetaURLGlobal = "https://graph.microsoft.com/beta/"
        $msGraphBetaURLUSGov = "https://graph.microsoft.us/beta/"
        $msGraphBetaURLUSDoD = "https://dod-graph.microsoft.us/beta/"
        $msGraphBetaURLChina = "https://microsoftgraph.chinacloudapi.cn/beta/"

        $msGraphGlobal = "Global"
        $msGraphUSGov = "USGov"
        $msGraphUSDOD = "USGovDOD"
        $msGraphChina = "China"

        $functionURL = ""

        Out-LogFile -string "********************************************************************************"
        Out-LogFile -string "BEGIN get-GraphEnvironment"
        Out-LogFile -string "********************************************************************************"
    
        Switch ($msGraphEnvironmentName)
        {
            $msGraphGlobal { out-logfile -string "Global URL" ; if ($useBeta -eq $TRUE) {$functionURL = $msGraphBetaURLGlobal} else {$functionURL = $msGraphURLGlobal}}
            $msGraphUSGov { out-logfile -string "USGovURL" ;  if ($useBeta -eq $TRUE) {$functionURL = $msGraphBetaURLUSGov} else {$functionURL = $msGraphURLUSGov}}
            $msGraphUSDOD { out-logfile -string "USDODUrl" ; if ($useBeta -eq $TRUE) {$functionURL = $msGraphBetaURLUSDoD} else {$functionURL = $msGraphURLUSDoD}}
            $msGraphChina { out-logfile -string "ChinaURL" ; if ($useBeta -eq $TRUE) {$functionURL = $msGraphBetaURLChina} else {$functionURL = $msGraphURLChina}}
        }

        out-logfile -string ("Graph URL Returned: "+$functionURL)

        Out-LogFile -string "END get-GraphEnvironment"
        Out-LogFile -string "********************************************************************************"

        return $functionURL
    }