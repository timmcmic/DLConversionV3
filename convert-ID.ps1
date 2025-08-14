function Convert-Id {
    <#
    .SYNOPSIS
    Convert ImmutableID to ObjectGuid OR ObjectGuid to ImmutableID
 
    .DESCRIPTION
    Convert ImmutableID to ObjectGuid OR ObjectGuid to ImmutableID
 
    .PARAMETER Id
    Enter an ImmutableID or ObjectGuid
 
    .EXAMPLE
    Convert-Id -Id 3MK05obeZEm9/xvs8svAFw==
 
    .EXAMPLE
    Convert-Id -Id e6b4csdc-des6-4d34-bdfe-9adff2cbc017
 
    .NOTES
    General notes
    #>

    param (
        [Parameter()]
        $Id
    )

    out-logfile -string $Id
    $Guid = [GUID]$Id
    out-logfile -string $guid
    $Byte = $Guid.ToByteArray()
    out-logfile -string $Byte
    $Object = [system.convert]::ToBase64String($Byte)

    return $Object
}
