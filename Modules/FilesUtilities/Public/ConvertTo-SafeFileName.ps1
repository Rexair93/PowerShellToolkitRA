function ConvertTo-SafeFileName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Name,

        [Parameter()]
        [string]$Replacement = '_'
    )

    $invalidChars = [IO.Path]::GetInvalidFileNameChars()
    $safeName = $Name

    foreach ($char in $invalidChars) {
        $safeName = $safeName.Replace($char, $Replacement)
    }

    $safeName = $safeName.Trim()

    if ([string]::IsNullOrWhiteSpace($safeName)) {
        return 'export'
    }

    return $safeName
}