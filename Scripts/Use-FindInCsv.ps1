Import-Module FilesUtilities -ErrorAction Stop

try {
    $csvRootPath = Get-FolderPath -Title 'Seleziona la cartella contenente i CSV da analizzare'

    $mode = Read-Host @"
Scegli modalità:
1 = ricerca stringa singola o elenco manuale
2 = ricerca termini ricavati dai nomi file di una cartella
"@

    $commonParams = @{
        CsvRootPath       = $csvRootPath
        Recurse           = $true
        IncludeSearchTerm = $true
    }

    $limitToColumns = Read-Host 'Vuoi limitare la ricerca a colonne specifiche? (S/N)'
    if ($limitToColumns -match '^[SsYy]') {
        $columnInput = Read-Host 'Inserisci i nomi colonna separati da virgola'
        $columns = $columnInput -split ',' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique

        if ($columns) {
            $commonParams.Column = $columns
        }
    }

    $exactChoice = Read-Host 'Ricerca esatta? (S/N)'
    if ($exactChoice -match '^[SsYy]') {
        $commonParams.ExactMatch = $true
    }

    $caseChoice = Read-Host 'Ricerca case-sensitive? (S/N)'
    if ($caseChoice -match '^[SsYy]') {
        $commonParams.CaseSensitive = $true
    }

    $searchTerms = @()

    switch ($mode) {
        '1' {
            $rawTerms = Read-Host 'Inserisci uno o più termini separati da virgola'
            $searchTerms = $rawTerms -split ',' |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
        }

        '2' {
            $searchFolderPath = Get-FolderPath -Title 'Seleziona la cartella da cui ricavare i termini di ricerca'

            $searchTerms = Get-ChildItem -Path $searchFolderPath -File |
                Select-Object -ExpandProperty BaseName |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
        }

        default {
            throw 'Modalità non valida.'
        }
    }

    if (-not $searchTerms) {
        throw 'Nessun termine di ricerca disponibile.'
    }

    $exportChoice = Read-Host 'Vuoi esportare i risultati? (S/N)'

    if ($exportChoice -match '^[SsYy]') {
        $exportFolder = Get-FolderPath -Title 'Seleziona la cartella in cui salvare gli export'

        foreach ($term in $searchTerms) {
            $results = Find-InCsv @commonParams -SearchValue $term

            if (-not $results) {
                Write-Host "Nessuna corrispondenza trovata per '$term'."
                continue
            }

            $safeFileName = ConvertTo-SafeFileName -Name $term

            $exportInfo = Get-ExportDestination `
                -DefaultFileName "$safeFileName.csv" `
                -InitialDirectory $exportFolder `
                -Formats @('csv', 'xlsx') `
                -PreferredFormat 'csv' `
                -Title "Scegli dove salvare i risultati per '$term'"

            $exported = Export-Results `
                -InputObject $results `
                -Path $exportInfo.Path `
                -Force

            Write-Host "Risultati per '$term' esportati in: $($exported.Path)"
        }
    }
    else {
        foreach ($term in $searchTerms) {
            $results = Find-InCsv @commonParams -SearchValue $term

            if (-not $results) {
                Write-Host "Nessuna corrispondenza trovata per '$term'."
                continue
            }

            Write-Host "`n=== Risultati per '$term' ==="
            $results | Format-Table -AutoSize
        }
    }
}
catch {
    Write-Error $_
}