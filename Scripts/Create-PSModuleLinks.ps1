[CmdletBinding()]
param (
    [string]$CustomModulesPath,
    [string]$LinksRootPath,

    [ValidateSet('SymbolicLink', 'Junction')]
    [string]$LinkType = 'SymbolicLink',

    [switch]$UseConsole
)

function Test-GuiAvailability {
    if (-not $IsWindows) {
        return $false
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}

function Get-FolderPath {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Title = 'Seleziona una cartella',

        [Parameter()]
        [string]$InitialDirectory = (Get-Location).Path,

        [Parameter()]
        [switch]$UseConsole,

        [Parameter()]
        [switch]$AllowEmpty
    )

    if (-not $UseConsole -and (Test-GuiAvailability)) {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Title
        $dialog.ShowNewFolderButton = $true

        if ($InitialDirectory -and (Test-Path $InitialDirectory -PathType Container)) {
            $dialog.SelectedPath = $InitialDirectory
        }

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }

        if ($AllowEmpty) {
            return $null
        }

        throw 'Operazione annullata.'
    }
    else {
        Write-Information 'Modalità console attiva.'

        $prompt = if ([string]::IsNullOrWhiteSpace($InitialDirectory)) {
            $Title
        }
        else {
            "$Title [$InitialDirectory]"
        }

        $path = Read-Host $prompt

        if ($null -ne $path) {
            $path = $path.Trim().Trim('"')
        }

        if ([string]::IsNullOrWhiteSpace($path)) {
            if ($AllowEmpty) {
                return $null
            }

            if (-not [string]::IsNullOrWhiteSpace($InitialDirectory)) {
                return $InitialDirectory
            }

            throw 'Percorso non specificato.'
        }

        return $path
    }
}

function Resolve-LinkTargetPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.IO.FileSystemInfo]$Item
    )

    if (-not $Item.LinkType) {
        return $null
    }

    $targetValue = $Item.Target

    if ($targetValue -is [array]) {
        $targetValue = $targetValue[0]
    }

    if ([string]::IsNullOrWhiteSpace($targetValue)) {
        return $null
    }

    try {
        return (Resolve-Path -LiteralPath $targetValue -ErrorAction Stop).Path
    }
    catch {
        try {
            $combined = Join-Path -Path $Item.DirectoryName -ChildPath $targetValue
            return (Resolve-Path -LiteralPath $combined -ErrorAction Stop).Path
        }
        catch {
            return $null
        }
    }
}

$defaultModulesRoot = Join-Path $HOME 'Documents\PowerShell\Modules'

#region Selezione cartella sorgente moduli
if (-not $CustomModulesPath) {
    $CustomModulesPath = Get-FolderPath `
        -Title 'Seleziona la cartella contenente i moduli PowerShell' `
        -UseConsole:$UseConsole
}

if (-not $CustomModulesPath) {
    Write-Warning 'Nessuna cartella sorgente selezionata. Operazione annullata.'
    return
}

if (-not (Test-Path -LiteralPath $CustomModulesPath -PathType Container)) {
    Write-Error "La cartella sorgente non esiste: $CustomModulesPath"
    return
}

$CustomModulesPath = (Resolve-Path -LiteralPath $CustomModulesPath).Path
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

if (-not (Test-Path -LiteralPath $LinksRootPath -PathType Container)) {
    New-Item -ItemType Directory -Path $LinksRootPath -Force | Out-Null
}

$LinksRootPath = (Resolve-Path -LiteralPath $LinksRootPath).Path
#endregion

#region Avviso SymbolicLink
if ($LinkType -eq 'SymbolicLink' -and $IsWindows) {
    Write-Verbose 'Su Windows la creazione di symbolic link può richiedere privilegi elevati o Developer Mode abilitato.'
}
#endregion

#region Elaborazione moduli
Get-ChildItem -LiteralPath $CustomModulesPath -Directory | ForEach-Object {
    $ModuleName = $_.Name
    $TargetPath = $_.FullName
    $LinkPath   = Join-Path $LinksRootPath $ModuleName

    if (Test-Path -LiteralPath $LinkPath) {
        $item = Get-Item -LiteralPath $LinkPath -Force -ErrorAction SilentlyContinue

        if ($item -and $item.LinkType) {
            $existingTarget = Resolve-LinkTargetPath -Item $item
            $currentTarget  = (Resolve-Path -LiteralPath $TargetPath).Path

            if ($existingTarget -and ($existingTarget -eq $currentTarget)) {
                Write-Host "✔ Link già presente per '$ModuleName' (corretto). Saltato." -ForegroundColor Green
                return
            }
            else {
                Write-Warning "⚠ '$ModuleName' esiste già ma NON è il link corretto. Saltato."
                return
            }
        }
        else {
            Write-Warning "⚠ '$ModuleName' esiste già come cartella/file normale in destinazione. Saltato."
            return
        }
    }

    try {
        $newItemSplat = @{
            ItemType = $LinkType
            Path     = $LinkPath
            Target   = $TargetPath
            ErrorAction = 'Stop'
        }

        if ($LinkType -eq 'Junction') {
            $newItemSplat.Target = (Resolve-Path -LiteralPath $TargetPath).Path
        }

        New-Item @newItemSplat | Out-Null
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