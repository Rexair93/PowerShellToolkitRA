function Export-Results {
    <#
    .SYNOPSIS
    Esporta una raccolta di oggetti in un file CSV, XLSX o ODS.

    .DESCRIPTION
    Gestisce tre formati:
      - csv  → Export-Csv nativo
      - xlsx → Export-Excel (modulo ImportExcel richiesto)
      - ods  → Export-Excel (ImportExcel >= 7.x supporta ODS in scrittura)

    Se il modulo ImportExcel non è disponibile e il formato richiesto è xlsx/ods,
    esegue un fallback a CSV avvertendo l'utente via Write-Warning.

    .PARAMETER InputObject
    Oggetti da esportare.

    .PARAMETER Path
    Percorso file di destinazione. L'estensione determina il formato.

    .PARAMETER WorksheetName
    Nome del foglio (solo xlsx/ods). Default: "Report".

    .PARAMETER Force
    Sovrascrive il file se già esiste (CSV). Per xlsx/ods la sovrascrittura
    è gestita sempre internamente da Export-Excel con -ClearSheet.

    .EXAMPLE
    $data | Export-Results -Path 'C:\out\report.xlsx' -WorksheetName 'BestOf'

    .EXAMPLE
    $data | Export-Results -Path 'C:\out\report.ods'

    .OUTPUTS
    PSCustomObject con proprietà Path e Format.
    #>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]] $InputObject,

        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter()]
        [string] $WorksheetName = 'Report',

        [Parameter()]
        [switch] $Force
    )

    begin {
        $collected = [System.Collections.Generic.List[object]]::new()
    }

    process {
        foreach ($item in $InputObject) {
            $collected.Add($item)
        }
    }

    end {
        $ext = ConvertTo-NormalizedExt ([IO.Path]::GetExtension($Path))
        $hasImportExcel = [bool](Get-Module -ListAvailable -Name ImportExcel)

        # --- XLSX / ODS ----------------------------------------------------------
        if ($ext -in 'xlsx', 'ods') {
            if ($hasImportExcel) {
                if ($PSCmdlet.ShouldProcess($Path, "Export $($ext.ToUpper())")) {
                    $dir = [IO.Path]::GetDirectoryName($Path)
                    if ($dir -and -not (Test-Path $dir)) {
                        New-Item -ItemType Directory -Path $dir -Force | Out-Null
                    }
                    $collected.ToArray() | Export-Excel -Path $Path `
                        -WorksheetName $WorksheetName -AutoSize -BoldTopRow -FreezeTopRow `
                        -ClearSheet -ErrorAction Stop
                }
                return [pscustomobject]@{ Path = $Path; Format = $ext }
            }

            # Fallback a CSV
            Write-Warning "Modulo ImportExcel non disponibile. Esportazione in CSV invece di '$ext'."
            $ext  = 'csv'
            $Path = [IO.Path]::ChangeExtension($Path, 'csv')
        }

        # --- CSV ------------------------------------------------------------------
        if ((Test-Path $Path) -and -not $Force) {
            throw "Il file '$Path' esiste già. Usa -Force per sovrascrivere."
        }

        if ($PSCmdlet.ShouldProcess($Path, 'Export CSV')) {
            $dir = [IO.Path]::GetDirectoryName($Path)
            if ($dir -and -not (Test-Path $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            $collected.ToArray() | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Force:$Force
        }

        return [pscustomobject]@{ Path = $Path; Format = $ext }
    }
}