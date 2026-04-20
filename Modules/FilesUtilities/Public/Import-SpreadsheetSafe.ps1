function Import-SpreadsheetSafe {
    <#
    .SYNOPSIS
    Importa un file tabellare (CSV, XLSX, ODS) restituendo oggetti PowerShell.

    .DESCRIPTION
    Interfaccia unificata di lettura per i formati supportati dal toolkit:
      - csv  → Import-CsvSafe (già nel modulo)
      - xlsx → Import-Excel   (richiede il modulo ImportExcel)
      - ods  → Import-Excel   (ImportExcel >= 7.x supporta ODS in lettura)

    Restituisce un array di PSCustomObject. In caso di file assente o vuoto
    restituisce un array vuoto senza eccezioni (comportamento analogo a Import-CsvSafe).

    .PARAMETER Path
    Percorso del file da importare.

    .PARAMETER Delimiter
    Delimitatore CSV. Default: virgola. Ignorato per xlsx/ods.

    .PARAMETER WorksheetName
    Nome del foglio da leggere (solo xlsx/ods). Se omesso viene letto il primo foglio.

    .EXAMPLE
    Import-SpreadsheetSafe -Path 'C:\data\report.xlsx'

    .EXAMPLE
    Import-SpreadsheetSafe -Path 'C:\data\report.csv' -Delimiter ';'

    .OUTPUTS
    PSCustomObject[]
    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$Path,

        [Parameter()]
        [char]$Delimiter = ',',

        [Parameter()]
        [string]$WorksheetName
    )

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        Write-Warning "File non trovato: '$Path'"
        return @()
    }

    $ext = ConvertTo-NormalizedExt ([IO.Path]::GetExtension($Path))

    switch ($ext) {

        'csv' {
            return Import-CsvSafe -Path $Path -Delimiter $Delimiter
        }

        { $_ -in 'xlsx', 'ods' } {
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                throw "Il modulo 'ImportExcel' non è disponibile. Installalo con: Install-Module ImportExcel"
            }

            $params = @{ Path = $Path; ErrorAction = 'Stop' }
            if ($WorksheetName) {
                $params['WorksheetName'] = $WorksheetName
            }

            $rows = Import-Excel @params
            if (-not $rows) { return @() }
            return $rows
        }

        default {
            throw "Formato '.$ext' non supportato da Import-SpreadsheetSafe. Formati ammessi: csv, xlsx, ods."
        }
    }
}