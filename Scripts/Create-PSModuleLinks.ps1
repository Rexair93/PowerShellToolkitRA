param (
    [string]$CustomModulesPath,
    
    [ValidateSet('SymbolicLink', 'Junction')]
    [string]$LinkType = 'SymbolicLink'
)

#region Selezione cartella (dialog o stringa)
if (-not $CustomModulesPath) {
    Add-Type -AssemblyName System.Windows.Forms

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Seleziona la cartella contenente i moduli PowerShell'
    $dialog.ShowNewFolderButton = $false

    if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Warning 'Nessuna cartella selezionata. Operazione annullata.'
        return
    }

    $CustomModulesPath = $dialog.SelectedPath
}
#endregion

#region Percorso moduli PowerShell
$PsModulesRoot = Join-Path $HOME 'Documents\PowerShell\Modules'

if (-not (Test-Path $PsModulesRoot)) {
    New-Item -ItemType Directory -Path $PsModulesRoot | Out-Null
}
#endregion

#region Elaborazione moduli
Get-ChildItem -Path $CustomModulesPath -Directory | ForEach-Object {

    $ModuleName = $_.Name
    $TargetPath = $_.FullName
    $LinkPath = Join-Path $PsModulesRoot $ModuleName

    if (Test-Path $LinkPath) {

        $item = Get-Item $LinkPath -ErrorAction SilentlyContinue

        # Controllo se è già un link e punta allo stesso target
        if ($item -and $item.LinkType -and
            (Resolve-Path $item.Target).Path -eq (Resolve-Path $TargetPath).Path) {

            Write-Host "✔ Link già presente per '$ModuleName' (corretto). Saltato." -ForegroundColor Green
            return
        }
        else {
            Write-Warning "⚠ '$ModuleName' esiste già ma NON è il link corretto. Saltato."
            return
        }
    }

    New-Item `
        -ItemType $LinkType `
        -Path $LinkPath `
        -Target $TargetPath | Out-Null

    Write-Host "➕ Creato link per modulo: $ModuleName" -ForegroundColor Cyan
}
#endregion

Write-Host "`nOperazione completata." -ForegroundColor Yellow