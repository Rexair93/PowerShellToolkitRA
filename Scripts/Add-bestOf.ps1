[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string]$InputPath,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$SearchValue,

    [Parameter()]
    [ValidateSet('csv', 'xlsx', 'ods')]
    [string]$ExportFormat,

    [Parameter()]
    [string]$OutputFileName = 'bestOf',

    [Parameter()]
    [char]$Delimiter = ',',

    [Parameter()]
    [switch]$Recurse,

    [Parameter()]
    [switch]$UseConsole,

    [Parameter()]
    [switch]$Force
)

Import-Module FilesUtilities -ErrorAction Stop

function Read-ChoiceValue {
    param(
        [Parameter(Mandatory)] [string]   $Prompt,
        [Parameter(Mandatory)] [string[]] $AllowedValues,
        [string] $Default
    )

    $allowed = $AllowedValues | ForEach-Object { $_.ToLowerInvariant() }

    while ($true) {
        $suffix = if ($Default) { " [$Default]" } else { '' }
        $raw = (Read-Host "$Prompt$suffix").Trim().ToLowerInvariant()
        $value = if ([string]::IsNullOrWhiteSpace($raw) -and $Default) { $Default.ToLowerInvariant() } else { $raw }

        if ($value -in $allowed) {
            return $value
        }

        Write-Warning "Valore non valido. Ammessi: $($AllowedValues -join ', ')"
    }
}

function Get-FileExtensionSafe {
    param([Parameter(Mandatory)][string]$Path)

    $ext = [IO.Path]::GetExtension($Path)
    if ([string]::IsNullOrWhiteSpace($ext)) {
        return $null
    }

    return $ext.TrimStart('.').ToLowerInvariant()
}

function Resolve-OutputTarget {
    param(
        [string]$OutputPath,
        [string]$OutputFileName,
        [string]$ExportFormat,
        [switch]$UseConsole,
        [switch]$Force
    )

    $safeName = ConvertTo-SafeFileName -Name $OutputFileName
    $allowedFormats = @('csv', 'xlsx', 'ods')

    if (-not $OutputPath) {
        if ($UseConsole) {
            $preferred = if ($ExportFormat) { $ExportFormat.ToLowerInvariant() } else { Read-ChoiceValue -Prompt 'Formato export' -AllowedValues $allowedFormats -Default 'csv' }
        }
        else {
            $preferred = if ($ExportFormat) { $ExportFormat.ToLowerInvariant() } else { $null }
        }

        return Get-ExportDestination -DefaultFileName "$safeName.csv" -Formats $allowedFormats -PreferredFormat $preferred -UseConsole:$UseConsole -Force:$Force
    }

    if (Test-Path $OutputPath -PathType Container) {
        $format = if ($ExportFormat) {
            $ExportFormat.ToLowerInvariant()
        }
        elseif ($UseConsole) {
            Read-ChoiceValue -Prompt 'Formato export' -AllowedValues $allowedFormats -Default 'csv'
        }
        else {
            throw "Se -OutputPath è una cartella e non usi -UseConsole, devi specificare anche -ExportFormat."
        }

        return [pscustomobject]@{
            Path   = (Join-Path $OutputPath "$safeName.$format")
            Format = $format
        }
    }

    $pathExt = Get-FileExtensionSafe -Path $OutputPath
    if ($pathExt) {
        if ($pathExt -notin $allowedFormats) {
            throw "Formato di output non supportato: '.$pathExt'. Formati ammessi: $($allowedFormats -join ', ')."
        }

        return [pscustomobject]@{
            Path   = $OutputPath
            Format = $pathExt
        }
    }

    $resolvedFormat = if ($ExportFormat) {
        $ExportFormat.ToLowerInvariant()
    }
    elseif ($UseConsole) {
        Read-ChoiceValue -Prompt 'Formato export' -AllowedValues $allowedFormats -Default 'csv'
    }
    else {
        throw "Se -OutputPath non ha estensione e non usi -UseConsole, devi specificare anche -ExportFormat."
    }

    return [pscustomobject]@{
        Path   = "$OutputPath.$resolvedFormat"
        Format = $resolvedFormat
    }
}

function Read-DuplicateAction {
    param([string]$FullName, [switch]$UseConsole)

    if ($UseConsole -or -not (Test-GuiAvailability)) {
        return Read-ChoiceValue -Prompt "FullName '$FullName' già presente. Scegli: overwrite / skip / duplicate" -AllowedValues @('overwrite', 'skip', 'duplicate') -Default 'skip'
    }

    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    $msg = "Il record con FullName:`n'$FullName'`nè già presente nel file di destinazione.`n`n• Sì      → Sovrascrivi il record esistente`n• No       → Inserisci come nuovo duplicato`n• Annulla → Salta questo record"
    $result = [System.Windows.Forms.MessageBox]::Show($msg, 'Duplicato rilevato', [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)

    switch ($result) {
        'Yes' { return 'overwrite' }
        'No' { return 'duplicate' }
        default { return 'skip' }
    }
}

function Read-RankingValue {
    while ($true) {
        $raw = Read-Host 'Inserisci Ranking (1-10), 0 o vuoto per nessun valore'
        if ([string]::IsNullOrWhiteSpace($raw) -or $raw -eq '0') {
            return $null
        }

        try {
            $n = [int]$raw
            if ($n -ge 1 -and $n -le 10) {
                return $n
            }
        }
        catch { }

        Write-Warning 'Valore non valido. Inserisci un intero da 1 a 10, 0 o vuoto.'
    }
}

if (-not $InputPath) {
    $InputPath = Get-FolderPath -Title 'Seleziona la cartella contenente i CSV' -UseConsole:$UseConsole
}

if (-not (Test-Path $InputPath -PathType Container)) {
    throw "Cartella input non trovata: '$InputPath'"
}

if (-not $SearchValue) {
    $SearchValue = (Read-Host 'Inserisci la stringa da cercare').Trim()
}

if (-not $Recurse) {
    $recurseChoice = Read-Host 'Includere sottocartelle? (S/N)'
    if ($recurseChoice -match '^[SsYy]') {
        $Recurse = $true
    }
}

if ([string]::IsNullOrWhiteSpace($SearchValue)) {
    throw 'Stringa di ricerca non valida.'
}

$destination = Resolve-OutputTarget -OutputPath $OutputPath -OutputFileName $OutputFileName -ExportFormat $ExportFormat -UseConsole:$UseConsole -Force:$Force
$outputFile = $destination.Path
$outputFormat = $destination.Format

# Mantiene l'ordine di inserimento e consente di aggiungere duplicati con una chiave interna unica.
$recordBuffer = [ordered]@{}

try {
    if (Test-Path $outputFile -PathType Leaf) {
        $existingRecords = Import-SpreadsheetSafe -Path $outputFile -Delimiter $Delimiter
        foreach ($r in $existingRecords) {
            $fullName = if ($r.PSObject.Properties['FullName']) { [string]$r.FullName } else { $null }
            $fullName = $fullName.Trim()
            if ([string]::IsNullOrWhiteSpace($fullName)) { continue }

            $recordKey = $fullName
            $suffix = 2
            while ($recordBuffer.Contains($recordKey)) {
                $recordKey = '{0} ({1})' -f $fullName, $suffix
                $suffix++
            }

            $recordBuffer[$recordKey] = $r
        }
        Write-Host "File esistente caricato: $outputFile ($($recordBuffer.Count) record unici, formato: $outputFormat)" -ForegroundColor Cyan
    }
} catch {
    Write-Error "Errore durante il caricamento del file di output esistente '$outputFile': $($_.Exception.Message)"
    # Continui l'esecuzione anche se non ci sono dati preesistenti, ma lo avvisa.
}

$found = Find-InCsv -CsvRootPath $InputPath -SearchValue $SearchValue -Recurse:$Recurse -IncludeSearchTerm -Delimiter $Delimiter
if (-not $found) {
    Write-Host 'Nessuna corrispondenza trovata.'
    return
}

foreach ($match in $found) {
    Write-Host "`n--- Record trovato in: $($match._SourceFile) ---" -ForegroundColor Green
    $match | Format-List

    $confirm = Read-ChoiceValue -Prompt 'Aggiungere questo record? (yes/no)' -AllowedValues @('yes', 'y', 'no', 'n') -Default 'no'
    if ($confirm -notin @('yes', 'y')) {
        Write-Host "Record saltato per scelta utente."
        continue
    }

    $ranking = Read-RankingValue

    $orderedRecord = [ordered]@{
        SourceFile = $match._SourceFile
        Ranking    = $ranking
    }

    foreach ($property in $match.PSObject.Properties) {
        if ($property.Name -in @('_SourceFile', '_SourcePath', '_SearchTerm', 'SourceFile', 'Ranking')) {
            continue
        }

        $orderedRecord[$property.Name] = $property.Value
    }

    $newRecord = [pscustomobject]$orderedRecord

    $fullName = if ($newRecord.PSObject.Properties['FullName']) { [string]$newRecord.FullName } else { $null }
    $fullName = if ($fullName) { $fullName.Trim() } else { $null }

    if ([string]::IsNullOrWhiteSpace($fullName)) {
        Write-Warning "Il record non contiene un FullName valido, viene saltato."
        continue
    }

    if ($recordBuffer.Contains($fullName)) {
        Write-Host "Duplicato rilevato: FullName '$fullName'." -ForegroundColor Yellow

        $action = Read-DuplicateAction -FullName $fullName -UseConsole:$UseConsole
        switch ($action) {
            'skip' {
                Write-Host 'Record saltato.' -ForegroundColor Gray
                continue
            }
            'overwrite' {
                $recordBuffer[$fullName] = $newRecord
                Write-Host "Record sovrascritto." -ForegroundColor Green
                continue
            }
            'duplicate' {
                $recordKey = $fullName
                $suffix = 2
                while ($recordBuffer.Contains($recordKey)) {
                    $recordKey = '{0} ({1})' -f $fullName, $suffix
                    $suffix++
                }

                $recordBuffer[$recordKey] = $newRecord
                Write-Host "Record duplicato aggiunto." -ForegroundColor Green
                continue
            }
        }
    }

    $recordBuffer[$fullName] = $newRecord
    Write-Host "Nuovo record aggiunto." -ForegroundColor Green
}

if ($recordBuffer.Count -eq 0) {
    Write-Host 'Nessun record da esportare.' -ForegroundColor Yellow
    return
}

# Convertire l'Hashtable dei record in un array per l'esportazione
$finalRecords = $recordBuffer.Values

if ($PSCmdlet.ShouldProcess($outputFile, 'Esportazione bestOf')) {
    try {
        $result = Export-Results -InputObject $finalRecords -Path $outputFile -WorksheetName 'bestOf' -Force
        Write-Host "==================================================" -ForegroundColor Cyan
        Write-Host "Esportazione completata con successo: $($result.Path) [formato: $($result.Format)]" -ForegroundColor Green
    } catch {
        Write-Error "Errore durante l'esportazione dei risultati in '$outputFile': $($_.Exception.Message)"
    }
} else {
    Write-Host 'Operazione di esportazione annullata.' -ForegroundColor Yellow
}
