function Find-InCsv {
    <#
    .SYNOPSIS
    Cerca una o più stringhe nei campi di file CSV presenti in una cartella.

    .DESCRIPTION
    Cerca uno o più termini nei valori dei campi di tutti i file CSV presenti in una cartella
    e, facoltativamente, nelle sottocartelle.

    Supporta due modalità:
    - ByValue: usa i termini passati tramite -SearchValue
    - ByFolder: ricava i termini dai nomi dei file presenti in una cartella

    I risultati vengono restituiti come oggetti PowerShell con colonne tecniche aggiuntive
    come _SourceFile, _SourcePath e _SearchTerm.

    .PARAMETER CsvRootPath
    Cartella contenente i CSV in cui effettuare la ricerca.

    .PARAMETER SearchValue
    Uno o più termini di ricerca.

    .PARAMETER SearchFolderPath
    Cartella da cui ricavare i termini di ricerca usando il BaseName dei file.

    .PARAMETER Filter
    Filtro file da usare per i CSV da analizzare. Default: *.csv

    .PARAMETER Column
    Una o più colonne specifiche in cui effettuare la ricerca. Se omesso, cerca in tutte.

    .PARAMETER Recurse
    Include le sottocartelle.

    .PARAMETER ExactMatch
    Esegue un confronto esatto.

    .PARAMETER CaseSensitive
    Esegue una ricerca case-sensitive.

    .PARAMETER IncludeSourcePath
    Include la colonna _SourcePath.

    .PARAMETER IncludeSearchTerm
    Include la colonna _SearchTerm.

    .PARAMETER IncludeCsvFile
    In modalità ByFolder, applica lo stesso filtro anche ai file usati per ricavare i termini.

    .PARAMETER Delimiter
    Delimitatore CSV. Default: virgola.

    .EXAMPLE
    Find-InCsv -CsvRootPath 'C:\Data' -SearchValue 'mario' -Recurse

    .EXAMPLE
    Find-InCsv -CsvRootPath 'C:\Data' -SearchFolderPath 'C:\Terms' -Recurse -IncludeSearchTerm

    .OUTPUTS
    PSCustomObject
    #>

    [CmdletBinding(DefaultParameterSetName = 'ByValue')]
    param(
        [Parameter(Mandatory, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [string]$CsvRootPath,

        [Parameter(Mandatory, ParameterSetName = 'ByValue', Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string[]]$SearchValue,

        [Parameter(Mandatory, ParameterSetName = 'ByFolder', Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$SearchFolderPath,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Filter = '*.csv',

        [Parameter()]
        [string[]]$Column,

        [Parameter()]
        [switch]$Recurse,

        [Parameter()]
        [switch]$ExactMatch,

        [Parameter()]
        [switch]$CaseSensitive,

        [Parameter()]
        [switch]$IncludeSourcePath,

        [Parameter()]
        [switch]$IncludeSearchTerm,

        [Parameter(ParameterSetName = 'ByFolder')]
        [switch]$IncludeCsvFile,

        [Parameter()]
        [char]$Delimiter = ','
    )

    begin {
        if (-not (Test-Path -Path $CsvRootPath -PathType Container)) {
            throw "Cartella CSV non trovata: '$CsvRootPath'"
        }

        $csvSearchParams = @{
            Path   = $CsvRootPath
            File   = $true
            Filter = $Filter
        }

        if ($Recurse) {
            $csvSearchParams.Recurse = $true
        }

        $csvFiles = Get-ChildItem @csvSearchParams | Sort-Object FullName

        if (-not $csvFiles) {
            throw "Nessun file trovato in '$CsvRootPath' con filtro '$Filter'."
        }

        $terms = switch ($PSCmdlet.ParameterSetName) {
            'ByValue' {
                $SearchValue |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique
            }

            'ByFolder' {
                if (-not (Test-Path -Path $SearchFolderPath -PathType Container)) {
                    throw "Cartella termini di ricerca non trovata: '$SearchFolderPath'"
                }

                $termSearchParams = @{
                    Path = $SearchFolderPath
                    File = $true
                }

                if ($Recurse) {
                    $termSearchParams.Recurse = $true
                }

                if ($IncludeCsvFile) {
                    $termSearchParams.Filter = $Filter
                }

                Get-ChildItem @termSearchParams |
                    Select-Object -ExpandProperty BaseName |
                    ForEach-Object { $_.Trim() } |
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    Sort-Object -Unique
            }
        }

        if (-not $terms) {
            throw 'Nessun termine di ricerca disponibile.'
        }

        $requestedColumns = @()
        if ($Column) {
            $requestedColumns = $Column |
                ForEach-Object { $_.Trim() } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
        }

        function Get-UniquePropertyName {
            param(
                [hashtable]$Map,
                [string]$Name
            )

            if (-not $Map.Contains($Name)) {
                return $Name
            }

            $index = 1
            do {
                $candidate = '{0}_{1}' -f $Name, $index
                $index++
            } until (-not $Map.Contains($candidate))

            return $candidate
        }
    }

    process {
        foreach ($term in $terms) {
            foreach ($file in $csvFiles) {
                try {
                    $rows = Import-CsvSafe -Path $file.FullName -Delimiter $Delimiter
                }
                catch {
                    Write-Warning "Impossibile importare '$($file.FullName)': $($_.Exception.Message)"
                    continue
                }

                foreach ($row in $rows) {
                    $propertiesToSearch = if ($requestedColumns) {
                        foreach ($columnName in $requestedColumns) {
                            $property = $row.PSObject.Properties[$columnName]
                            if ($null -ne $property) {
                                $property
                            }
                        }
                    }
                    else {
                        $row.PSObject.Properties
                    }

                    if (-not $propertiesToSearch) {
                        continue
                    }

                    $isMatch = $false

                    foreach ($property in $propertiesToSearch) {
                        $value = [string]$property.Value

                        if ($CaseSensitive) {
                            if ($ExactMatch) {
                                if ($value -ceq $term) {
                                    $isMatch = $true
                                    break
                                }
                            }
                            else {
                                if ($value -clike "*$term*") {
                                    $isMatch = $true
                                    break
                                }
                            }
                        }
                        else {
                            if ($ExactMatch) {
                                if ($value -eq $term) {
                                    $isMatch = $true
                                    break
                                }
                            }
                            else {
                                if ($value -like "*$term*") {
                                    $isMatch = $true
                                    break
                                }
                            }
                        }
                    }

                    if (-not $isMatch) {
                        continue
                    }

                    $orderedProperties = [ordered]@{}

                    if ($IncludeSearchTerm) {
                        $orderedProperties['_SearchTerm'] = $term
                    }

                    $orderedProperties['_SourceFile'] = $file.Name

                    if ($IncludeSourcePath) {
                        $orderedProperties['_SourcePath'] = $file.FullName
                    }

                    foreach ($property in $row.PSObject.Properties) {
                        $propertyName = [string]$property.Name

                        if ([string]::IsNullOrWhiteSpace($propertyName)) {
                            $propertyName = 'UnnamedColumn'
                        }

                        $safeName = Get-UniquePropertyName -Map $orderedProperties -Name $propertyName
                        $orderedProperties[$safeName] = $property.Value
                    }

                    [pscustomobject]$orderedProperties
                }
            }
        }
    }
}