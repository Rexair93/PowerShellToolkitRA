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
            [string[]] $searchTerms = $rawTerms -split ',' |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique

            if (-not $searchTerms) {
                throw 'Nessun termine di ricerca disponibile.'
            }

            $exportChoice = Read-Host 'Vuoi esportare i risultati? (S/N)'

            if ($exportChoice -match '^[SsYy]') {
                if ($searchTerms.Count -eq 1) {
                    $term = $searchTerms[0]
                    $results = Find-InCsv @commonParams -SearchValue $term

                    if (-not $results) {
                        Write-Host "Nessuna corrispondenza trovata per '$term'."
                        return
                    }

                    $safeFileName = ConvertTo-SafeFileName -Name $term

                    $exportInfo = Get-ExportDestination `
                        -DefaultFileName "$safeFileName.csv" `
                        -InitialDirectory $csvRootPath `
                        -Formats @('csv', 'xlsx') `
                        -PreferredFormat 'csv' `
                        -Title "Salva risultati per '$term'"

                    $exported = Export-Results `
                        -InputObject $results `
                        -Path $exportInfo.Path `
                        -Force

                    Write-Host "Risultati per '$term' esportati in: $($exported.Path)"
                }
                else {
                    $exportFolder = Get-FolderPath -Title 'Seleziona la cartella in cui salvare gli export multipli'
                    $formatChoice = Read-Host 'Formato export multiplo (csv/xlsx) [csv]'

                    if ([string]::IsNullOrWhiteSpace($formatChoice)) {
                        $formatChoice = 'csv'
                    }

                    $formatChoice = $formatChoice.Trim().ToLowerInvariant()

                    if ($formatChoice -notin @('csv', 'xlsx')) {
                        throw "Formato non valido: '$formatChoice'"
                    }

                    foreach ($term in $searchTerms) {
                        $results = Find-InCsv @commonParams -SearchValue $term

                        if (-not $results) {
                            Write-Host "Nessuna corrispondenza trovata per '$term'."
                            continue
                        }

                        $safeFileName = ConvertTo-SafeFileName -Name $term
                        $outputPath = Join-Path $exportFolder "$safeFileName.$formatChoice"

                        $exported = Export-Results `
                            -InputObject $results `
                            -Path $outputPath `
                            -Force

                        Write-Host "Risultati per '$term' esportati in: $($exported.Path)"
                    }
                }
            }
            else {
                foreach ($term in $searchTerms) {
                    $results = Find-InCsv @commonParams -SearchValue $term

                    if ($results) {
                        Write-Host "`n=== Risultati per '$term' ==="
                        $results | Format-Table -AutoSize
                    }
                    else {
                        Write-Host "Nessuna corrispondenza per '$term'."
                    }
                }
            }
        }

        '2' {
            $searchFolderPath = Get-FolderPath -Title 'Seleziona la cartella da cui ricavare i termini di ricerca'
            $exportChoice = Read-Host 'Vuoi esportare i risultati? (S/N)'

            $searchTerms = Get-ChildItem -Path $searchFolderPath -File |
            Select-Object -ExpandProperty BaseName |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique

            if (-not $searchTerms) {
                throw 'Nessun termine di ricerca disponibile.'
            }

            if ($exportChoice -match '^[SsYy]') {
                $exportFolder = Get-FolderPath -Title 'Seleziona la cartella in cui salvare gli export automatici'
                $formatChoice = Read-Host 'Formato export automatico (csv/xlsx) [csv]'

                if ([string]::IsNullOrWhiteSpace($formatChoice)) {
                    $formatChoice = 'csv'
                }

                $formatChoice = $formatChoice.Trim().ToLowerInvariant()

                if ($formatChoice -notin @('csv', 'xlsx')) {
                    throw "Formato non valido: '$formatChoice'"
                }

                foreach ($term in $searchTerms) {
                    $results = Find-InCsv @commonParams -SearchValue $term

                    if (-not $results) {
                        Write-Host "Nessuna corrispondenza trovata per '$term'."
                        continue
                    }

                    $safeFileName = ConvertTo-SafeFileName -Name $term
                    $outputPath = Join-Path $exportFolder "$safeFileName.$formatChoice"

                    $exported = Export-Results `
                        -InputObject $results `
                        -Path $outputPath `
                        -Force

                    Write-Host "Risultati per '$term' esportati automaticamente in: $($exported.Path)"
                }
            }
            else {
                foreach ($term in $searchTerms) {
                    $results = Find-InCsv @commonParams -SearchValue $term

                    if ($results) {
                        Write-Host "`n=== Risultati per '$term' ==="
                        $results | Format-Table -AutoSize
                    }
                    else {
                        Write-Host "Nessuna corrispondenza per '$term'."
                    }
                }
            }
        }

        default {
            throw 'Modalità non valida.'
        }
    }
}
catch {
    Write-Error $_
}