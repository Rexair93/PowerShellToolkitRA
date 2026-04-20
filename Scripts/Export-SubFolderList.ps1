#Requires -Version 7.0
#Requires -Modules FilesUtilities

<#
.SYNOPSIS
    Esporta la lista delle sottocartelle di primo livello in un file txt, csv o xlsx.

.DESCRIPTION
    Chiede all'utente la cartella di analisi e la destinazione di export,
    quindi chiama Get-SubFolderList (modulo FilesUtilities) ed esporta i risultati.
    In modalità GUI il formato viene scelto dalla finestra di dialogo.
    In modalità console (-UseConsole) il formato viene chiesto interattivamente
    oppure può essere passato direttamente tramite -ExportFormat.

.PARAMETER FolderPath
    Percorso della cartella principale da analizzare.
    Se non specificato, viene richiesto tramite Get-FolderPath.

.PARAMETER OutputPath
    Percorso completo del file di output.
    Se non specificato, viene richiesto tramite Get-ExportDestination.

.PARAMETER ExportFormat
    Formato di export: txt (predefinito), csv, xlsx.
    Rilevante solo in modalità console (-UseConsole).
    Se non specificato, viene chiesto interattivamente.

.PARAMETER UseConsole
    Forza la selezione dei percorsi in modalità console (nessuna finestra di dialogo).

.PARAMETER Force
    Sovrascrive il file di output senza chiedere conferma.

.EXAMPLE
    .\Invoke-SubFolderExport.ps1

.EXAMPLE
    .\Invoke-SubFolderExport.ps1 -UseConsole

.EXAMPLE
    .\Invoke-SubFolderExport.ps1 -FolderPath "C:\Progetti" -ExportFormat csv -UseConsole

.EXAMPLE
    .\Invoke-SubFolderExport.ps1 -FolderPath "C:\Progetti" -OutputPath "C:\Report\sottocartelle.xlsx"
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string] $FolderPath,

    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [ValidateSet('txt', 'csv', 'xlsx')]
    [string] $ExportFormat,

    [Parameter()]
    [switch] $UseConsole,

    [Parameter()]
    [switch] $Force
)

Import-Module FilesUtilities -ErrorAction Stop

# ── 1. Cartella sorgente ─────────────────────────────────────────────────────
if (-not $FolderPath) {
    $FolderPath = Get-FolderPath `
        -Title "Seleziona la cartella principale da analizzare" `
        -UseConsole:$UseConsole
}

if (-not $FolderPath) {
    Write-Warning "Nessuna cartella selezionata. Operazione annullata."
    return
}

# ── 2. Raccolta dati ─────────────────────────────────────────────────────────
$subfolders = Get-SubFolderList -FolderPath $FolderPath

if (-not $subfolders) {
    Write-Warning "Nessuna sottocartella trovata in '$FolderPath'."
    return
}

Write-Host "`nTrovate $($subfolders.Count) sottocartelle in '$FolderPath'." -ForegroundColor Cyan

# ── 3. Scelta formato (solo in modalità console, se non già specificato) ─────
$guiMode = -not $UseConsole -and (Test-GuiAvailability)

if (-not $guiMode -and -not $ExportFormat) {
    Write-Host "`nIn quale formato vuoi esportare il risultato?"
    Write-Host "  [1] txt  (predefinito)"
    Write-Host "  [2] csv"
    Write-Host "  [3] xlsx"
    $formatChoice = Read-Host "`nScelta [1/2/3]"

    $ExportFormat = switch ($formatChoice.Trim()) {
        '2' { 'csv' }
        '3' { 'xlsx' }
        default { 'txt' }
    }
}

# In modalità GUI il formato viene gestito direttamente dalla finestra di dialogo.
# ExportFormat (se passato) viene comunque propagato come PreferredFormat.
$defaultName = Split-Path $FolderPath -Leaf
$availableFormats = @('txt', 'csv', 'xlsx')
$preferredFmt = if ($ExportFormat) { $ExportFormat } else { 'txt' }

# ── 4. Destinazione export ───────────────────────────────────────────────────
if (-not $OutputPath) {
    $dest = Get-ExportDestination `
        -DefaultFileName "$defaultName.$preferredFmt" `
        -Formats $availableFormats `
        -PreferredFormat $preferredFmt `
        -Title "Scegli dove salvare il file" `
        -UseConsole:$UseConsole `
        -Force:$Force
}
else {
    $ext = [IO.Path]::GetExtension($OutputPath).TrimStart('.').ToLowerInvariant()
    if (-not $ext) {
        $OutputPath = "$OutputPath.$preferredFmt"
        $ext = $preferredFmt
    }
    $dest = [pscustomobject]@{ Path = $OutputPath; Format = $ext }
}

# ── 5. Export ────────────────────────────────────────────────────────────────
switch ($dest.Format) {
    'txt' {
        if ($PSCmdlet.ShouldProcess($dest.Path, "Export TXT")) {
            $subfolders.FullName | Out-File -FilePath $dest.Path -Encoding UTF8 -Force:$Force
        }
        Write-Host "`nFile TXT salvato in: $($dest.Path)" -ForegroundColor Green
    }
    default {
        $result = Export-Results `
            -InputObject $subfolders `
            -Path $dest.Path `
            -WorksheetName "Sottocartelle" `
            -Force:$Force
        Write-Host "`nFile $($result.Format.ToUpper()) salvato in: $($result.Path)" -ForegroundColor Green
    }
}