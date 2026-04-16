function Import-CsvSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [char]$Delimiter = ','
    )

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        throw "File non trovato: '$Path'"
    }

    $rawLines = Get-Content -Path $Path
    if (-not $rawLines) {
        return @()
    }

    if ($rawLines.Count -lt 2) {
        return @()
    }

    $headerLine = $rawLines[0]
    $dataLines  = $rawLines | Select-Object -Skip 1

    $headerValues = $headerLine -split [regex]::Escape([string]$Delimiter)

    $uniqueHeaders = [System.Collections.Generic.List[string]]::new()
    $seenHeaders = @{}

    foreach ($header in $headerValues) {
        $name = [string]$header
        $name = $name.Trim().Trim('"')

        if ([string]::IsNullOrWhiteSpace($name)) {
            $name = 'UnnamedColumn'
        }

        if ($seenHeaders.ContainsKey($name)) {
            $seenHeaders[$name]++
            $name = '{0}_{1}' -f $name, $seenHeaders[$name]
        }
        else {
            $seenHeaders[$name] = 0
        }

        $uniqueHeaders.Add($name)
    }

    $csvBody = $dataLines -join [Environment]::NewLine

    if ([string]::IsNullOrWhiteSpace($csvBody)) {
        return @()
    }

    return $csvBody | ConvertFrom-Csv -Delimiter $Delimiter -Header $uniqueHeaders
}