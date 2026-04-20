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

$buffer = [System.Collections.Generic.List[object]]::new()
if (Test-Path $outputFile -PathType Leaf) {
    $existing = Import-SpreadsheetSafe -Path $outputFile -Delimiter $Delimiter
    foreach ($r in $existing) {
        [void]$buffer.Add($r)
    }
    Write-Host "File esistente caricato: $outputFile ($($buffer.Count) record, formato: $outputFormat)"
}

$found = Find-InCsv -CsvRootPath $InputPath -SearchValue $SearchValue -Recurse:$Recurse -IncludeSearchTerm -Delimiter $Delimiter
if (-not $found) {
    Write-Host 'Nessuna corrispondenza trovata.'
    return
}

foreach ($match in $found) {
    Write-Host "`n--- Record trovato in: $($match._SourceFile) ---"
    $match | Format-List

    $confirm = Read-ChoiceValue -Prompt 'Aggiungere questo record? (yes/no)' -AllowedValues @('yes', 'y', 'no', 'n') -Default 'no'
    if ($confirm -notin @('yes', 'y')) {
        continue
    }

    $ranking = Read-RankingValue

    $ordered = [ordered]@{}
    $ordered['SourceFile'] = $match._SourceFile
    $ordered['Ranking'] = $ranking

    foreach ($prop in $match.PSObject.Properties) {
        if ($prop.Name -in '_SourceFile', '_SourcePath', '_SearchTerm') {
            continue
        }
        $ordered[$prop.Name] = $prop.Value
    }

    $newRecord = [pscustomobject]$ordered
    $fullName = if ($newRecord.PSObject.Properties['FullName']) { [string]$newRecord.FullName } else { $null }

    if (-not [string]::IsNullOrWhiteSpace($fullName)) {
        $dupIdx = @(for ($i = 0; $i -lt $buffer.Count; $i++) {
                $existing = $buffer[$i]
                $fp = $existing.PSObject.Properties['FullName']
                if ($fp -and [string]$fp.Value -eq $fullName) { $i }
            })

        if ($dupIdx.Count -gt 0) {
            $action = Read-DuplicateAction -FullName $fullName -UseConsole:$UseConsole
            switch ($action) {
                'skip' {
                    Write-Host 'Record saltato.'
                    continue
                }
                'overwrite' {
                    foreach ($idx in ($dupIdx | Sort-Object -Descending)) {
                        $buffer.RemoveAt($idx)
                    }
                }
            }
        }
    }

    [void]$buffer.Add($newRecord)
}

if ($buffer.Count -eq 0) {
    Write-Host 'Nessun record da esportare.'
    return
}

if ($PSCmdlet.ShouldProcess($outputFile, 'Esportazione bestOf')) {
    $result = Export-Results -InputObject $buffer.ToArray() -Path $outputFile -WorksheetName 'bestOf' -Force
    Write-Host "Esportazione completata: $($result.Path) [formato: $($result.Format)]"
}