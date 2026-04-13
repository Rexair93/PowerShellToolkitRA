function Export-Results {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [object[]] $InputObject,

        [Parameter(Mandatory)]
        [string] $Path,

        [Parameter()]
        [string] $WorksheetName = "Report",

        [Parameter()]
        [switch] $Force
    )

    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()

    if ($ext -eq ".xlsx" -and (Get-Module -ListAvailable -Name ImportExcel)) {
        if ($PSCmdlet.ShouldProcess($Path, "Export Excel")) {
            $InputObject | Export-Excel -Path $Path `
                -WorksheetName $WorksheetName -AutoSize -BoldTopRow -FreezeTopRow `
                -ClearSheet -ErrorAction Stop
        }
        return [pscustomobject]@{ Path = $Path; Format = "xlsx" }
    }

    # fallback CSV (anche se l'utente aveva richiesto xlsx)
    $csvPath = if ($ext -eq ".xlsx") { [IO.Path]::ChangeExtension($Path, ".csv") } else { $Path }

    if ((Test-Path $csvPath) -and -not $Force) {
        throw "Il file '$csvPath' esiste già. Usa -Force per sovrascrivere."
    }

    if ($PSCmdlet.ShouldProcess($csvPath, "Export CSV")) {
        $InputObject | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8 -Force:$Force
    }

    return [pscustomobject]@{ Path = $csvPath; Format = "csv" }
}