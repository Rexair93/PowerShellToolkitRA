function Get-FolderContentReport {
    <#
    .SYNOPSIS
        Restituisce un report del contenuto di una o più cartelle, includendo opzionalmente metadati multimediali.

    .DESCRIPTION
        Elenca file e cartelle presenti nei percorsi specificati e produce un oggetto con le proprietà principali
        del filesystem. Per impostazione predefinita include anche una serie di tag multimediali letti tramite
        Windows Shell Property System, con possibilità di disabilitarli o personalizzarli.

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
        @{ NomeColonnaOutput = 'Nome Proprietà Shell' }

        Esempio:
        @{ Performer = 'Artists'; AlbumTitle = 'Album'; PublishedYear = 'Year' }

    .EXAMPLE
        Get-FolderContentReport -FolderPath 'D:\Media'

    .EXAMPLE
        Get-FolderContentReport -FolderPath 'D:\Media' -IncludeMediaTags $false

    .EXAMPLE
        Get-FolderContentReport -FolderPath 'D:\Media' -Recurse -MediaTagNames Artists,Album,Title,Genre

    .EXAMPLE
        Get-FolderContentReport -FolderPath 'D:\Media' -Recurse `
            -MediaTagNames Artists,Album,Title `
            -CustomMediaTags @{ Performer = 'Artists'; TrackNo = 'TrackNumber' }

    .NOTES
        La disponibilità dei metadati dipende dalle proprietà esposte dalla Shell di Windows
        per il tipo di file specifico.
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
            'Artist',
            'Title',
            'Duration',
            'Year',
            'Releasetime',
            'Distributor',
            'Publisher',
            'Studio',
            'Site',
            'Bitrate',
            'Rating'
        ),

        [hashtable] $CustomMediaTags
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

        function Get-ShellPropertyMap {
            param(
                [Parameter(Mandatory)]
                [object] $ShellFolder
            )

            $propertyMap = @{}

            for ($index = 0; $index -le 400; $index++) {
                $propertyName = $ShellFolder.GetDetailsOf($null, $index)

                if (-not [string]::IsNullOrWhiteSpace($propertyName)) {
                    $normalizedPropertyName = ($propertyName -replace '\s+', '').Trim().ToLowerInvariant()

                    if (-not $propertyMap.ContainsKey($normalizedPropertyName)) {
                        $propertyMap[$normalizedPropertyName] = $index
                    }
                }
            }

            return $propertyMap
        }

        function Resolve-ShellPropertyIndex {
            param(
                [Parameter(Mandatory)]
                [hashtable] $PropertyMap,

                [Parameter(Mandatory)]
                [string] $PropertyName
            )

            $normalizedPropertyName = ($PropertyName -replace '\s+', '').Trim().ToLowerInvariant()

            if ($PropertyMap.ContainsKey($normalizedPropertyName)) {
                return $PropertyMap[$normalizedPropertyName]
            }

            return $null
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
                    $shellPropertyName = [string] $RequestedTags[$outputName]
                    $propertyIndex = Resolve-ShellPropertyIndex -PropertyMap $propertyMap -PropertyName $shellPropertyName

                    if ($null -eq $propertyIndex) {
                        continue
                    }

                    $rawValue = $shellFolder.GetDetailsOf($shellItem, $propertyIndex)

                    if (-not [string]::IsNullOrWhiteSpace($rawValue)) {
                        $values[$outputName] = $rawValue.Trim()
                    }
                }
            }
            catch {
                Write-Verbose "Impossibile leggere i metadati media per '$($File.FullName)': $($_.Exception.Message)"
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
                    $shellPropertyName = [string] $CustomMediaTags[$customKey]

                    if ([string]::IsNullOrWhiteSpace($outputName)) {
                        continue
                    }

                    if ([string]::IsNullOrWhiteSpace($shellPropertyName)) {
                        continue
                    }

                    $resolvedMediaTagMap[$outputName] = $shellPropertyName
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

                if ($IncludeMediaTags -and $resolvedMediaTagMap.Count -gt 0) {
                    $mediaTagValues = Get-ShellMediaTagValues -File $entry -RequestedTags $resolvedMediaTagMap -ShellObject $shell -FolderCache $shellFolderCache

                    foreach ($mediaKey in $mediaTagValues.Keys) {
                        $itemData[$mediaKey] = $mediaTagValues[$mediaKey]
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