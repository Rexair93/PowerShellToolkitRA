<#
.SYNOPSIS
    Sincronizza i file PowerShell (.ps1, .psm1, .psd1) da una cartella sorgente
    verso una cartella destinazione come file .txt, e genera struttura.txt.

.DESCRIPTION
    Per default apre finestre di dialogo Windows per la selezione delle cartelle.
    Con -UseConsole forza l'inserimento da riga di comando.
    I parametri -SourceFolder e -DestinationFolder permettono di bypassare
    qualsiasi richiesta interattiva (utili per automazione/scheduling).

.PARAMETER SourceFolder
    Cartella sorgente contenente i file PowerShell.
    Se omesso, viene aperta una finestra di dialogo (o richiesta console).

.PARAMETER DestinationFolder
    Cartella destinazione dove copiare i file .txt e struttura.txt.
    Se omessa, viene aperta una finestra di dialogo (o richiesta console).

.PARAMETER UseConsole
    Se specificato, forza la modalità console anche su sistemi con GUI disponibile.

.EXAMPLE
    # Selezione tramite finestre di dialogo (default su Windows con GUI)
    .\Sync-PSFilesToTxt.ps1

.EXAMPLE
    # Selezione tramite console (no GUI)
    .\Sync-PSFilesToTxt.ps1 -UseConsole

.EXAMPLE
    # Percorsi passati direttamente via parametro (nessuna finestra/prompt)
    .\Sync-PSFilesToTxt.ps1 -SourceFolder "C:\Scripts" -DestinationFolder "C:\Output"
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(HelpMessage = "Cartella sorgente con i file PowerShell")]
    [string]$SourceFolder,

    [Parameter(HelpMessage = "Cartella destinazione per i file .txt e struttura.txt")]
    [string]$DestinationFolder,

    [Parameter(HelpMessage = "Forza la modalità console (nessuna GUI)")]
    [switch]$UseConsole,

    [Parameter(HelpMessage = "Nome della cartella radice da cui far partire il percorso in struttura.txt (es. 'PowerShellToolkitRA')")]
    [string]$DisplayRootFolder = 'PowerShellToolkitRA'
)

# ---------------------------------------------------------------------------
# Costanti
# ---------------------------------------------------------------------------
$PS_EXTENSIONS  = @('.ps1', '.psm1', '.psd1')
$STRUTTURA_FILE = 'struttura.txt'

# ---------------------------------------------------------------------------
# Funzione: verifica disponibilità GUI Windows
# (stessa logica di Test-GuiAvailability nel modulo FileUtilities)
# ---------------------------------------------------------------------------
function Test-GuiAvailability {
    if (-not $IsWindows -and $PSVersionTable.PSVersion.Major -ge 6) {
        return $false
    }
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

# ---------------------------------------------------------------------------
# Funzione: apre FolderBrowserDialog o chiede il percorso da console
# ---------------------------------------------------------------------------
function Select-Folder {
    param (
        [string]$Title       = "Seleziona cartella",
        [string]$Description = "",
        [bool]  $ForceConsole = $false
    )

    if (-not $ForceConsole -and (Test-GuiAvailability)) {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $dlg = [System.Windows.Forms.FolderBrowserDialog]::new()
        $dlg.Description         = if ($Description) { $Description } else { $Title }
        $dlg.UseDescriptionForTitle = $true   # mostra il testo nella barra del titolo (Windows Vista+)
        $dlg.ShowNewFolderButton = $true
        $dlg.RootFolder          = [System.Environment+SpecialFolder]::MyComputer

        $result = $dlg.ShowDialog()
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            throw "Operazione annullata dall'utente."
        }
        return $dlg.SelectedPath
    } else {
        Write-Host "$Title" -ForegroundColor Cyan
        $path = (Read-Host "Inserisci il percorso della cartella").Trim()
        if ([string]::IsNullOrWhiteSpace($path)) {
            throw "Percorso non specificato."
        }
        return $path
    }
}

# ---------------------------------------------------------------------------
# Funzione: Genera struttura ad albero della cartella sorgente
# ---------------------------------------------------------------------------
function Get-FolderTree {
    param (
        [string]$Path,
        [string]$Indent = '',
        [bool]  $IsLast = $true
    )

    $lines       = [System.Collections.Generic.List[string]]::new()
    $branch      = if ($IsLast) { '└── ' } else { '├── ' }
    $dirName     = Split-Path -Leaf $Path
    $lines.Add("$Indent$branch$dirName")

    $childSuffix = if ($IsLast) { '    ' } else { '│   ' }
    $childIndent = $Indent + $childSuffix

    $children = Get-ChildItem -LiteralPath $Path | Sort-Object { $_.PSIsContainer } -Descending |
                Sort-Object Name

    for ($i = 0; $i -lt $children.Count; $i++) {
        $child     = $children[$i]
        $childLast = ($i -eq $children.Count - 1)

        if ($child.PSIsContainer) {
            $subLines = Get-FolderTree -Path $child.FullName -Indent $childIndent -IsLast $childLast
            foreach ($line in $subLines) { $lines.Add([string]$line) }
        } else {
            $childBranch = if ($childLast) { '└── ' } else { '├── ' }
            $lines.Add("$childIndent$childBranch$($child.Name)")
        }
    }

    return $lines
}

# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

# --- Risolvi cartella sorgente ---
if ([string]::IsNullOrWhiteSpace($SourceFolder)) {
    $SourceFolder = Select-Folder `
        -Title       "Seleziona la cartella SORGENTE" `
        -Description "Seleziona la cartella contenente i file PowerShell (.ps1, .psm1, .psd1)" `
        -ForceConsole $UseConsole.IsPresent
}

# Valida che la cartella sorgente esista
if (-not (Test-Path $SourceFolder -PathType Container)) {
    Write-Error "Cartella sorgente non trovata: '$SourceFolder'"
    exit 1
}
$SourceFolder = (Resolve-Path $SourceFolder).Path

# --- Risolvi cartella destinazione ---
if ([string]::IsNullOrWhiteSpace($DestinationFolder)) {
    $DestinationFolder = Select-Folder `
        -Title       "Seleziona la cartella DESTINAZIONE" `
        -Description "Seleziona (o crea) la cartella dove salvare i file .txt e struttura.txt" `
        -ForceConsole $UseConsole.IsPresent
}

$DestinationFolder = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DestinationFolder)

Write-Verbose "Sorgente     : $SourceFolder"
Write-Verbose "Destinazione : $DestinationFolder"

# Crea la cartella destinazione se non esiste
if (-not (Test-Path $DestinationFolder)) {
    if ($PSCmdlet.ShouldProcess($DestinationFolder, "Crea cartella destinazione")) {
        New-Item -Path $DestinationFolder -ItemType Directory -Force | Out-Null
        Write-Verbose "Creata cartella destinazione: $DestinationFolder"
    }
}

# ---------------------------------------------------------------------------
# 1. Sincronizza i file PowerShell come .txt (tutti nella root di destinazione)
# ---------------------------------------------------------------------------
$psFiles = Get-ChildItem -Path $SourceFolder -Recurse -File |
           Where-Object { $PS_EXTENSIONS -contains $_.Extension.ToLower() }

if (-not $psFiles) {
    Write-Warning "Nessun file .ps1, .psm1 o .psd1 trovato in: $SourceFolder"
} else {
    $copied  = 0
    $updated = 0
    $skipped = 0

    foreach ($file in $psFiles) {
        # Tutti i file finiscono nella root della destinazione (nessuna sottocartella)
        $destPath = Join-Path $DestinationFolder ($file.Name + '.txt')

        $shouldCopy = $false
        $action     = ''

        if (-not (Test-Path $destPath)) {
            $shouldCopy = $true
            $action     = 'Creato'
            $copied++
        } elseif ($file.LastWriteTimeUtc -gt (Get-Item $destPath).LastWriteTimeUtc) {
            $shouldCopy = $true
            $action     = 'Aggiornato'
            $updated++
        } else {
            $skipped++
            Write-Verbose "Saltato (invariato): $($file.Name)"
        }

        if ($shouldCopy) {
            if ($PSCmdlet.ShouldProcess($destPath, $action)) {
                Copy-Item -LiteralPath $file.FullName -Destination $destPath -Force
                $color = if ($action -eq 'Creato') { 'Green' } else { 'Cyan' }
                Write-Host "$action : $($file.FullName) → $destPath" -ForegroundColor $color
            }
        }
    }

    Write-Host "`nRiepilogo sincronizzazione:" -ForegroundColor Yellow
    Write-Host "  Nuovi file creati  : $copied"  -ForegroundColor Green
    Write-Host "  File aggiornati    : $updated" -ForegroundColor Cyan
    Write-Host "  File invariati     : $skipped" -ForegroundColor Gray
}

# ---------------------------------------------------------------------------
# 2. Genera / aggiorna struttura.txt nella cartella destinazione
# ---------------------------------------------------------------------------
$strutturaDest = Join-Path $DestinationFolder $STRUTTURA_FILE

if ($PSCmdlet.ShouldProcess($strutturaDest, "Genera struttura.txt")) {
    $alreadyExists = Test-Path $strutturaDest

    # Calcola il percorso da mostrare in struttura.txt
    $displayPath = if (-not [string]::IsNullOrWhiteSpace($DisplayRootFolder)) {
        # Trova la prima occorrenza del nome radice nel percorso (case-insensitive)
        $idx = $SourceFolder.IndexOf($DisplayRootFolder, [System.StringComparison]::OrdinalIgnoreCase)
        if ($idx -ge 0) {
            $SourceFolder.Substring($idx)
        } else {
            Write-Warning "DisplayRootFolder '$DisplayRootFolder' non trovato nel percorso sorgente. Verra' usato il percorso completo."
            $SourceFolder
        }
    } else {
        # Fallback: rimuove la home utente se il percorso vi si trova sotto
        $homePath = [System.Environment]::GetFolderPath('UserProfile')
        if ($SourceFolder.StartsWith($homePath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $SourceFolder.Substring($homePath.Length).TrimStart('\', '/')
        } else {
            $SourceFolder
        }
    }

    $header = @(
        "Struttura della cartella sorgente",
        "Generato il : $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')",
        "Sorgente    : $displayPath",
        ("─" * 60)
    )

    # Prima riga = nome root
    $rootName  = Split-Path $SourceFolder -Leaf
    $treeLines = [System.Collections.Generic.List[string]]::new()
    $treeLines.Add($rootName)

    $children = Get-ChildItem -LiteralPath $SourceFolder |
                Sort-Object { $_.PSIsContainer } -Descending |
                Sort-Object Name

    for ($i = 0; $i -lt $children.Count; $i++) {
        $child  = $children[$i]
        $isLast = ($i -eq $children.Count - 1)

        if ($child.PSIsContainer) {
            $subLines = Get-FolderTree -Path $child.FullName -Indent '' -IsLast $isLast
            foreach ($line in $subLines) { $treeLines.Add([string]$line) }
        } else {
            $branch = if ($isLast) { '└── ' } else { '├── ' }
            $treeLines.Add("$branch$($child.Name)")
        }
    }

    $content = ($header + $treeLines) -join "`n"
    Set-Content -Path $strutturaDest -Value $content -Encoding UTF8

    $verb = if ($alreadyExists) { "Aggiornato" } else { "Creato" }
    Write-Host "`n$verb struttura.txt in: $strutturaDest" -ForegroundColor Magenta
}

Write-Host "`nOperazione completata." -ForegroundColor Yellow