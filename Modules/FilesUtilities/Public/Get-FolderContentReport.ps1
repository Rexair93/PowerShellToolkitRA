function Get-FolderContentReport {
    <#
    .SYNOPSIS
        Restituisce un report del contenuto di una o più cartelle, includendo metadati multimediali standard
        e, opzionalmente, tag custom tramite ExifTool.

    .DESCRIPTION
        Elenca file e cartelle presenti nei percorsi specificati e produce un oggetto con le proprietà principali
        del filesystem. Per impostazione predefinita include anche una serie di tag multimediali standard letti
        tramite Windows Shell Property System.

        Se vengono definiti CustomMediaTags e ExifTool è disponibile, la funzione tenta inoltre di leggere
        i tag custom embedded nel file (ad esempio quelli creati con Mp3tag) e li aggiunge al report.

    .PARAMETER FolderPath
        Uno o più percorsi di cartelle da analizzare.

    .PARAMETER Extensions
        Estensioni file da includere nel report. Le estensioni vengono normalizzate.

    .PARAMETER IncludeSourceFolder
        Include nell'output la cartella sorgente risolta.

    .PARAMETER Recurse
        Analizza ricorsivamente le sottocartelle.

    .PARAMETER IncludeMediaTags
        Indica se includere i metadati multimediali nel report.
        Il valore predefinito è $true.

    .PARAMETER MediaTagNames
        Elenco di nomi proprietà Shell da esportare come colonne aggiuntive.
        Se non specificato, viene usata una lista di default.

    .PARAMETER CustomMediaTags
        Hashtable di alias personalizzati nel formato:
        @{ NomeColonnaOutput = 'NomeTagOrigine' }

        Per i tag standard/Shell il valore può essere il nome proprietà di Windows.
        Per i tag custom il valore deve essere il nome leggibile da ExifTool, ad esempio:
        @{ Mood = 'TXXX:MOOD'; CatalogNo = 'TXXX:CATALOGNUMBER' }

    .PARAMETER ExifToolPath
        Percorso esplicito di exiftool.exe. Se omesso, la funzione prova a trovarlo nel PATH.

    .EXAMPLE
        Get-FolderContentReport -FolderPath 'D:\Media' -Recurse

    .EXAMPLE
        Get-FolderContentReport -FolderPath 'D:\Media' -Recurse `
            -CustomMediaTags @{
                Mood      = 'TXXX:MOOD'
                CatalogNo = 'TXXX:CATALOGNUMBER'
            }

    .NOTES
        I nomi delle proprietà Shell possono essere localizzati.
        I tag custom Mp3tag richiedono ExifTool per essere letti in modo affidabile.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'Path')]
        [string[]] $FolderPath,

        [string[]] $Extensions,

        [switch] $IncludeSourceFolder,

        [switch] $Recurse,

        [bool] $IncludeMediaTags = $true,

        [string[]] $MediaTagNames = @(
            'Titolo',
            'Artisti partecipanti',
            'Durata',
            'Anno'
        ),

        [hashtable] $CustomMediaTags = @{
            Title = 'ItemList:Title'
            Publisher = 'ItemList:Publisher'
            Artist = 'ItemList:Artist'
            Director = 'ItemList:Director'
            Studio = 'iTunes:STUDIO'
            Site = 'iTunes:SITE'
            ReleaseTime = 'iTunes:RELEASETIME'
            Distributor = 'iTunes:DISTRIBUTOR'
        },

        [string] $ExifToolPath = 'C:\exiftool\exiftool.exe'
    )

    begin {
        $results = New-Object System.Collections.Generic.List[object]

        $normalizedExtensions = @()
        if ($Extensions) {
            $normalizedExtensions = $Extensions |
                ConvertTo-NormalizedExt |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique
        }

        $hasExtensionFilter = $normalizedExtensions.Count -gt 0

        $baseProperties = @(
            'Name',
            'Size [MB]',
            'Extension',
            'CreationTime',
            'LastAccessTime',
            'LastWriteTime',
            'FullName',
            'Length',
            'BaseName',
            'Directory',
            'PSIsContainer'
        )

        $propertiesWithSourceFolder = @('SourceFolder') + $baseProperties

        $shell = $null
        $shellFolderCache = @{}
        $resolvedMediaTagMap = [ordered]@{}
        $resolvedCustomTagMap = [ordered]@{}
        $effectiveExifToolPath = $null
        $canUseExifToolForCustomTags = $false

        function Normalize-MetadataName {
            param(
                [Parameter(Mandatory)]
                [string] $Name
            )

            $normalized = $Name.Normalize([Text.NormalizationForm]::FormD)
            $normalized = [regex]::Replace($normalized, '\p{Mn}', '')
            $normalized = $normalized.ToLowerInvariant()
            $normalized = $normalized -replace '[^a-z0-9:]', ''

            return $normalized
        }

        function Get-ShellPropertyMap {
            param(
                [Parameter(Mandatory)]
                [object] $ShellFolder
            )

            $propertyMap = @{}

            for ($index = 0; $index -le 400; $index++) {
                $propertyName = $ShellFolder.GetDetailsOf($null, $index)

                if (-not [string]::IsNullOrWhiteSpace($propertyName)) {
                    $normalizedPropertyName = Normalize-MetadataName -Name $propertyName

                    if (-not $propertyMap.ContainsKey($normalizedPropertyName)) {
                        $propertyMap[$normalizedPropertyName] = @{
                            Index = $index
                            Name  = $propertyName
                        }
                    }
                }
            }

            return $propertyMap
        }

        function Resolve-ShellProperty {
            param(
                [Parameter(Mandatory)]
                [hashtable] $PropertyMap,

                [Parameter(Mandatory)]
                [string] $RequestedName
            )

            $candidateNames = @($RequestedName)

            switch -Regex (Normalize-MetadataName -Name $RequestedName) {
                '^artists?$' {
                    $candidateNames += @('Artisti partecipanti', 'Artisti', 'Autori', 'Contributing artists', 'Participating artists')
                    break
                }
                '^title$' {
                    $candidateNames += @('Titolo', 'Nome')
                    break
                }
                '^album$' {
                    $candidateNames += @('Nome album')
                    break
                }
                '^genre$' {
                    $candidateNames += @('Genere')
                    break
                }
                '^year$' {
                    $candidateNames += @('Anno')
                    break
                }
                '^tracknumber$' {
                    $candidateNames += @('Numero traccia', 'Track number')
                    break
                }
                '^duration$' {
                    $candidateNames += @('Durata', 'Lunghezza')
                    break
                }
                '^bitrate$' {
                    $candidateNames += @('Velocità in bit', 'Bit rate')
                    break
                }
            }

            foreach ($candidateName in ($candidateNames | Select-Object -Unique)) {
                $normalizedCandidate = Normalize-MetadataName -Name $candidateName

                if ($PropertyMap.ContainsKey($normalizedCandidate)) {
                    return $PropertyMap[$normalizedCandidate]
                }
            }

            return $null
        }

        function Resolve-ExifToolExecutable {
            param(
                [string] $PreferredPath
            )

            if (-not [string]::IsNullOrWhiteSpace($PreferredPath)) {
                if (Test-Path -Path $PreferredPath -PathType Leaf) {
                    return (Resolve-Path -Path $PreferredPath).Path
                }

                Write-Warning "ExifTool non trovato nel percorso specificato: '$PreferredPath'."
                return $null
            }

            $command = Get-Command -Name 'exiftool.exe' -ErrorAction SilentlyContinue
            if ($null -ne $command) {
                return $command.Source
            }

            $command = Get-Command -Name 'exiftool' -ErrorAction SilentlyContinue
            if ($null -ne $command) {
                return $command.Source
            }

            return $null
        }

        function Get-ExifToolCustomTagValues {
            param(
                [Parameter(Mandatory)]
                [System.IO.FileSystemInfo] $File,

                [Parameter(Mandatory)]
                [string] $ToolPath,

                [Parameter(Mandatory)]
                [hashtable] $RequestedTags
            )

            $values = [ordered]@{}

            foreach ($outputName in $RequestedTags.Keys) {
                $values[$outputName] = $null
            }

            if ($File.PSIsContainer) {
                return $values
            }

            $requestedSourceNames = $RequestedTags.Values | Select-Object -Unique
            $arguments = @('-j', '-G1', '-a', '-s')

            foreach ($sourceName in $requestedSourceNames) {
                $arguments += "-$sourceName"
            }

            $arguments += $File.FullName

            try {
                $jsonOutput = & $ToolPath @arguments 2>$null
                $jsonText = $jsonOutput -join [Environment]::NewLine

                if ([string]::IsNullOrWhiteSpace($jsonText)) {
                    return $values
                }

                $parsedOutput = $jsonText | ConvertFrom-Json
                if ($parsedOutput -is [array]) {
                    $parsedOutput = $parsedOutput | Select-Object -First 1
                }

                if ($null -eq $parsedOutput) {
                    return $values
                }

                $propertyBag = @{}
                foreach ($property in $parsedOutput.PSObject.Properties) {
                    $propertyBag[$property.Name] = $property.Value
                }

                foreach ($outputName in $RequestedTags.Keys) {
                    $sourceName = [string] $RequestedTags[$outputName]
                    $normalizedSourceName = Normalize-MetadataName -Name $sourceName

                    foreach ($propertyName in $propertyBag.Keys) {
                        $normalizedPropertyName = Normalize-MetadataName -Name $propertyName

                        if ($normalizedPropertyName -eq $normalizedSourceName -or
                            $normalizedPropertyName.EndsWith(":$normalizedSourceName")) {
                            $rawValue = $propertyBag[$propertyName]

                            if ($rawValue -is [array]) {
                                $values[$outputName] = ($rawValue -join '; ')
                            }
                            elseif ($null -ne $rawValue -and -not [string]::IsNullOrWhiteSpace([string] $rawValue)) {
                                $values[$outputName] = [string] $rawValue
                            }

                            break
                        }
                    }
                }
            }
            catch {
                Write-Verbose "ExifTool non è riuscito a leggere i tag custom per '$($File.FullName)': $($_.Exception.Message)"
            }

            return $values
        }

        function Get-ShellMediaTagValues {
            param(
                [Parameter(Mandatory)]
                [System.IO.FileSystemInfo] $File,

                [Parameter(Mandatory)]
                [hashtable] $RequestedTags,

                [Parameter(Mandatory)]
                [object] $ShellObject,

                [Parameter(Mandatory)]
                [hashtable] $FolderCache
            )

            $values = [ordered]@{}

            foreach ($outputName in $RequestedTags.Keys) {
                $values[$outputName] = $null
            }

            if ($File.PSIsContainer) {
                return $values
            }

            if ([string]::IsNullOrWhiteSpace($File.DirectoryName)) {
                return $values
            }

            try {
                if (-not $FolderCache.ContainsKey($File.DirectoryName)) {
                    $shellFolder = $ShellObject.Namespace($File.DirectoryName)

                    if ($null -eq $shellFolder) {
                        return $values
                    }

                    $FolderCache[$File.DirectoryName] = @{
                        Folder      = $shellFolder
                        PropertyMap = (Get-ShellPropertyMap -ShellFolder $shellFolder)
                    }
                }

                $folderInfo = $FolderCache[$File.DirectoryName]
                $shellFolder = $folderInfo.Folder
                $propertyMap = $folderInfo.PropertyMap
                $shellItem = $shellFolder.ParseName($File.Name)

                if ($null -eq $shellItem) {
                    return $values
                }

                foreach ($outputName in $RequestedTags.Keys) {
                    $resolvedProperty = Resolve-ShellProperty -PropertyMap $propertyMap -RequestedName ([string] $RequestedTags[$outputName])

                    if ($null -eq $resolvedProperty) {
                        continue
                    }

                    $rawValue = $shellFolder.GetDetailsOf($shellItem, $resolvedProperty.Index)

                    if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
                        $values[$outputName] = $rawValue.Trim()
                    }
                }
            }
            catch {
                Write-Verbose "Impossibile leggere i metadati Shell per '$($File.FullName)': $($_.Exception.Message)"
            }

            return $values
        }

        if ($IncludeMediaTags) {
            try {
                $shell = New-Object -ComObject Shell.Application
            }
            catch {
                throw "Impossibile inizializzare Shell.Application per la lettura dei metadati multimediali. Dettagli: $($_.Exception.Message)"
            }

            foreach ($tagName in ($MediaTagNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
                $resolvedMediaTagMap[$tagName] = $tagName
            }

            if ($CustomMediaTags) {
                foreach ($customKey in $CustomMediaTags.Keys) {
                    $outputName = [string] $customKey
                    $sourceName = [string] $CustomMediaTags[$customKey]

                    if ([string]::IsNullOrWhiteSpace($outputName) -or [string]::IsNullOrWhiteSpace($sourceName)) {
                        continue
                    }

                    $resolvedCustomTagMap[$outputName] = $sourceName
                }
            }

            if ($resolvedCustomTagMap.Count -gt 0) {
                $effectiveExifToolPath = Resolve-ExifToolExecutable -PreferredPath $ExifToolPath

                if ([string]::IsNullOrWhiteSpace($effectiveExifToolPath)) {
                    Write-Verbose "ExifTool non trovato. I tag custom non verranno popolati."
                }
                else {
                    $canUseExifToolForCustomTags = $true
                    Write-Verbose "ExifTool rilevato: $effectiveExifToolPath"
                }
            }
        }
    }

    process {
        foreach ($currentFolder in $FolderPath) {
            if ([string]::IsNullOrWhiteSpace($currentFolder)) {
                Write-Warning "Percorso cartella vuoto o non valido. Saltato."
                continue
            }

            if (-not (Test-Path -Path $currentFolder -PathType Container)) {
                Write-Warning "La cartella '$currentFolder' non esiste o non è accessibile. Saltata."
                continue
            }

            $resolvedFolder = (Resolve-Path -Path $currentFolder).Path
            $folderName = Split-Path -Path $resolvedFolder -Leaf

            $getChildItemParams = @{
                Path  = $resolvedFolder
                Force = $true
            }

            if ($Recurse) {
                $getChildItemParams.Recurse = $true
            }

            if ($hasExtensionFilter) {
                $getChildItemParams.File = $true
            }

            $items = Get-ChildItem @getChildItemParams

            if ($hasExtensionFilter) {
                $items = $items | Where-Object {
                    (ConvertTo-NormalizedExt $_.Extension) -in $normalizedExtensions
                }
            }

            foreach ($entry in $items) {
                $itemData = [ordered]@{
                    SourceFolder   = $folderName
                    Name           = $entry.Name
                    'Size [MB]'    = if ($entry.PSIsContainer -or $null -eq $entry.Length) {
                        $null
                    }
                    else {
                        [Math]::Round(($entry.Length / 1MB), 2)
                    }
                    Extension      = $entry.Extension
                    CreationTime   = $entry.CreationTime
                    LastAccessTime = $entry.LastAccessTime
                    LastWriteTime  = $entry.LastWriteTime
                    FullName       = $entry.FullName
                    Length         = $entry.Length
                    BaseName       = $entry.BaseName
                    Directory      = $entry.DirectoryName
                    PSIsContainer  = $entry.PSIsContainer
                }

                if ($IncludeMediaTags -and -not $entry.PSIsContainer) {
                    $shellTagValues = Get-ShellMediaTagValues -File $entry -RequestedTags $resolvedMediaTagMap -ShellObject $shell -FolderCache $shellFolderCache

                    foreach ($tagName in $resolvedMediaTagMap.Keys) {
                        $itemData[$tagName] = $shellTagValues[$tagName]
                    }

                    if ($canUseExifToolForCustomTags) {
                        $customTagValues = Get-ExifToolCustomTagValues -File $entry -ToolPath $effectiveExifToolPath -RequestedTags $resolvedCustomTagMap

                        foreach ($customTagName in $resolvedCustomTagMap.Keys) {
                            $itemData[$customTagName] = $customTagValues[$customTagName]
                        }
                    }
                    elseif ($resolvedCustomTagMap.Count -gt 0) {
                        foreach ($customTagName in $resolvedCustomTagMap.Keys) {
                            $itemData[$customTagName] = $null
                        }
                    }
                }

                $item = [pscustomobject] $itemData

                $selectedProperties = if ($IncludeSourceFolder) {
                    @($propertiesWithSourceFolder)
                }
                else {
                    @($baseProperties)
                }

                if ($IncludeMediaTags -and $resolvedMediaTagMap.Count -gt 0) {
                    $selectedProperties += @($resolvedMediaTagMap.Keys)
                }

                if ($resolvedCustomTagMap.Count -gt 0) {
                    $selectedProperties += @($resolvedCustomTagMap.Keys)
                }

                $results.Add(($item | Select-Object -Property $selectedProperties))
            }
        }
    }

    end {
        if ($null -ne $shell) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
        }

        $results.ToArray()
    }
}