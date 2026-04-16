[CmdletBinding()]
param (
    [string]$CustomModulesPath,
    [string]$LinksRootPath,

    [ValidateSet('SymbolicLink', 'Junction')]
    [string]$LinkType = 'SymbolicLink',

    [switch]$UseConsole
)

Import-Module FilesUtilities -ErrorAction Stop

$defaultModulesRoot = Join-Path $HOME 'Documents\PowerShell\Modules'

#region Selezione cartella sorgente moduli
if (-not $CustomModulesPath) {
    $CustomModulesPath = Get-FolderPath `
        -Title 'Seleziona la cartella contenente i moduli PowerShell' `
        -UseConsole:$UseConsole

    if (-not $CustomModulesPath) {
        Write-Warning 'Nessuna cartella sorgente selezionata. Operazione annullata.'
        return
    }
}

if (-not (Test-Path $CustomModulesPath -PathType Container)) {
    Write-Error "La cartella sorgente non esiste: $CustomModulesPath"
    return
}
#endregion

#region Selezione cartella destinazione link
if (-not $LinksRootPath) {
    $LinksRootPath = Get-FolderPath `
        -Title 'Seleziona la cartella in cui creare i symlink dei moduli' `
        -InitialDirectory $defaultModulesRoot `
        -UseConsole:$UseConsole `
        -AllowEmpty

    if (-not $LinksRootPath) {
        Write-Host "Nessuna cartella destinazione selezionata. Uso il percorso predefinito: $defaultModulesRoot" -ForegroundColor Yellow
        $LinksRootPath = $defaultModulesRoot
    }
}

if (-not (Test-Path $LinksRootPath -PathType Container)) {
    New-Item -ItemType Directory -Path $LinksRootPath -Force | Out-Null
}
#endregion

#region Elaborazione moduli
Get-ChildItem -Path $CustomModulesPath -Directory | ForEach-Object {
    $ModuleName = $_.Name
    $TargetPath = $_.FullName
    $LinkPath   = Join-Path $LinksRootPath $ModuleName

    if (Test-Path $LinkPath) {
        $item = Get-Item $LinkPath -Force -ErrorAction SilentlyContinue

        if ($item -and $item.LinkType) {
            try {
                $existingTarget = (Resolve-Path $item.Target -ErrorAction Stop).Path
                $currentTarget  = (Resolve-Path $TargetPath -ErrorAction Stop).Path

                if ($existingTarget -eq $currentTarget) {
                    Write-Host "✔ Link già presente per '$ModuleName' (corretto). Saltato." -ForegroundColor Green
                    return
                }
                else {
                    Write-Warning "⚠ '$ModuleName' esiste già ma punta a un'altra destinazione. Saltato."
                    return
                }
            }
            catch {
                Write-Warning "⚠ '$ModuleName' esiste come link, ma non è stato possibile verificarne il target. Saltato."
                return
            }
        }
        else {
            Write-Warning "⚠ '$ModuleName' esiste già come cartella/file normale in destinazione. Saltato."
            return
        }
    }

    try {
        New-Item `
            -ItemType $LinkType `
            -Path $LinkPath `
            -Target $TargetPath `
            -ErrorAction Stop | Out-Null

        Write-Host "➕ Creato link per modulo: $ModuleName" -ForegroundColor Cyan
    }
    catch {
        Write-Warning "❌ Errore creando il link per '$ModuleName': $($_.Exception.Message)"
    }
}
#endregion

Write-Host "`nOperazione completata." -ForegroundColor Yellow
Write-Host "Sorgente moduli : $CustomModulesPath" -ForegroundColor DarkGray
Write-Host "Destinazione link: $LinksRootPath" -ForegroundColor DarkGray
Write-Host "Tipo link       : $LinkType" -ForegroundColor DarkGray
Write-Host "Modalità input  : $(if ($UseConsole) { 'Console' } else { 'Automatica (GUI se disponibile, altrimenti console)' })" -ForegroundColor DarkGray