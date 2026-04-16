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

    switch ($mode) {
        '1' {
            $rawTerms = Read-Host 'Inserisci uno o più termini separati da virgola'
            $searchValues = $rawTerms -split ',' |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique

            if (-not $searchValues) {
                throw 'Nessun termine di ricerca specificato.'
            }

            $results = Find-InCsv @commonParams -SearchValue $searchValues
        }

        '2' {
            $searchFolderPath = Get-FolderPath -Title 'Seleziona la cartella da cui ricavare i termini di ricerca'
            $results = Find-InCsv @commonParams -SearchFolderPath $searchFolderPath
        }

        default {
            throw 'Modalità non valida.'
        }
    }

    if (-not $results) {
        Write-Host 'Nessuna corrispondenza trovata.'
        return
    }

    Write-Host "Trovate $($results.Count) corrispondenze."
    $results | Format-Table -AutoSize

    $exportChoice = Read-Host 'Vuoi esportare i risultati? (S/N)'
    if ($exportChoice -match '^[SsYy]') {
        $export = Get-ExportDestination `
            -DefaultFileName 'Find-InCsv_Report.csv' `
            -InitialDirectory $csvRootPath `
            -Formats @('csv', 'xlsx') `
            -PreferredFormat 'csv' `
            -Title 'Scegli il file di destinazione'

        $exported = Export-Results `
            -InputObject $results `
            -Path $export.Path `
            -Force

        Write-Host "Risultati esportati in: $($exported.Path)"
    }
}
catch {
    Write-Error $_
}