[CmdletBinding()]
param(
    [Parameter()]
    [ValidateSet('Single', 'List')]
    [string] $Mode,

    [Parameter()]
    [ValidateSet('csv', 'xlsx')]
    [string] $ExportFormat,

    [Parameter()]
    [switch] $UseConsole
)

Import-Module FilesUtilities -ErrorAction Stop

function Get-MultiExportFormat {
    [CmdletBinding()]
    param(
        [string] $Default = 'csv'
    )

    $choice = Read-Host "Formato export per i file multipli (csv/xlsx) [$Default]"
    if ([string]::IsNullOrWhiteSpace($choice)) {
        return $Default
    }

    $normalized = $choice.Trim().ToLowerInvariant()
    if ($normalized -notin @('csv', 'xlsx')) {
        throw "Formato non valido: '$choice'. Valori ammessi: csv, xlsx."
    }

    return $normalized
}

function Export-SingleFolderContent {
    [CmdletBinding()]
    param(
        [string] $ExportFormat,
        [switch] $UseConsole
    )

    $inputPath = Get-FolderPath `
        -Title "Seleziona la cartella di cui esportare il contenuto" `
        -AllowEmpty `
        -UseConsole:$UseConsole

    if (-not $inputPath) {
        Write-Warning "Nessuna cartella selezionata. Operazione annullata."
        return
    }

    if (-not (Test-Path -Path $inputPath -PathType Container)) {
        throw "La cartella '$inputPath' non esiste o non è valida."
    }

    $resolvedInputPath = (Resolve-Path -Path $inputPath).Path
    $folderName = Split-Path -Path $resolvedInputPath -Leaf

    $destinationParams = @{
        DefaultFileName  = "$folderName.csv"
        InitialDirectory = (Get-Location).Path
        Formats          = @('csv', 'xlsx')
        Title            = "Scegli dove salvare l'export della cartella"
        UseConsole       = $UseConsole
        Force            = $true
    }

    if ($ExportFormat) {
        $destinationParams.DefaultFileName = "$folderName.$ExportFormat"
        $destinationParams.PreferredFormat = $ExportFormat
    }

    $destination = Get-ExportDestination @destinationParams
    $data = Get-FolderContentReport -FolderPath $resolvedInputPath
    $result = Export-Results -InputObject $data -Path $destination.Path -WorksheetName $folderName -Force

    Write-Host ""
    Write-Host "Esportazione completata:" -ForegroundColor Green
    Write-Host $result.Path
    Write-Host ""
}

function Export-MultipleFolderContents {
    [CmdletBinding()]
    param(
        [string] $ExportFormat,
        [switch] $UseConsole
    )

    $inputListFile = Get-InputFile `
        -Formats @('txt') `
        -Title "Seleziona il file TXT con la lista delle cartelle" `
        -UseConsole:$UseConsole

    if (-not $inputListFile) {
        Write-Warning "Nessun file di input selezionato. Operazione annullata."
        return
    }

    $outputFolder = Get-FolderPath `
        -Title "Seleziona la cartella in cui esportare i file" `
        -AllowEmpty `
        -UseConsole:$UseConsole

    if (-not $outputFolder) {
        Write-Warning "Nessuna cartella di destinazione selezionata. Operazione annullata."
        return
    }

    if (-not (Test-Path -Path $outputFolder -PathType Container)) {
        throw "La cartella di destinazione '$outputFolder' non esiste o non è valida."
    }

    $selectedFormat = if ($ExportFormat) {
        $ExportFormat
    }
    else {
        Get-MultiExportFormat -Default 'csv'
    }

    $folderPaths = Get-Content -Path $inputListFile |
    ForEach-Object { $_.Trim() } |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    if (-not $folderPaths) {
        throw "Il file '$inputListFile' non contiene percorsi validi."
    }

    foreach ($inputPath in $folderPaths) {
        if (-not (Test-Path -Path $inputPath -PathType Container)) {
            Write-Warning "Il path '$inputPath' non esiste o non è una cartella. Saltato."
            continue
        }

        try {
            $resolvedInputPath = (Resolve-Path -Path $inputPath).Path
            $folderName = Split-Path -Path $resolvedInputPath -Leaf
            $outputFilePath = Join-Path -Path $outputFolder -ChildPath "$folderName.$selectedFormat"

            $data = Get-FolderContentReport -FolderPath $resolvedInputPath
            $result = Export-Results -InputObject $data -Path $outputFilePath -WorksheetName $folderName -Force

            Write-Host ""
            Write-Host "Esportata lista per '$resolvedInputPath'" -ForegroundColor Green
            Write-Host "File creato: $($result.Path)"
            Write-Host ""
        }
        catch {
            Write-Warning "Errore durante l'elaborazione di '$inputPath': $($_.Exception.Message)"
        }
    }
}

if (-not $Mode) {
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host " Export Folder Content" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Seleziona modalità:"
    Write-Host "1 - Cartella singola"
    Write-Host "2 - Elenco cartelle da file TXT"
    Write-Host ""

    $choice = Read-Host "Inserisci la scelta (1 o 2)"

    switch ($choice) {
        '1' { $Mode = 'Single' }
        '2' { $Mode = 'List' }
        default { throw "Scelta non valida. Inserire 1 oppure 2." }
    }
}

switch ($Mode) {
    'Single' { Export-SingleFolderContent -ExportFormat $ExportFormat -UseConsole:$UseConsole }
    'List' { Export-MultipleFolderContents -ExportFormat $ExportFormat -UseConsole:$UseConsole }
}