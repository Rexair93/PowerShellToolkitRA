function Get-FolderContentReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'Path')]
        [string[]] $FolderPath,

        [string[]] $Extensions,

        [switch] $IncludeSourceFolder,

        [switch] $Recurse
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
                $item = [pscustomobject]@{
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

                if ($IncludeSourceFolder) {
                    $results.Add(($item | Select-Object -Property $propertiesWithSourceFolder))
                }
                else {
                    $results.Add(($item | Select-Object -Property $baseProperties))
                }
            }
        }
    }

    end {
        $results.ToArray()
    }
}