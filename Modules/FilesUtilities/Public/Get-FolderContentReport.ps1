function Get-FolderContentReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'Path')]
        [string[]] $FolderPath,

        [switch] $IncludeSourceFolder
    )

    begin {
        $results = New-Object System.Collections.Generic.List[object]
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

            Get-ChildItem -Path $resolvedFolder -Force | ForEach-Object {
                $item = [pscustomobject]@{
                    SourceFolder   = $folderName
                    Name           = $_.Name
                    'Size [MB]'    = [Math]::Round(($_.Length / 1MB), 2)
                    Extension      = $_.Extension
                    CreationTime   = $_.CreationTime
                    LastAccessTime = $_.LastAccessTime
                    LastWriteTime  = $_.LastWriteTime
                    FullName       = $_.FullName
                    Length         = $_.Length
                    BaseName       = $_.BaseName
                    Directory      = $_.DirectoryName
                    PSIsContainer  = $_.PSIsContainer
                }

                if (-not $IncludeSourceFolder) {
                    $item = $item | Select-Object Name, 'Size [MB]', Extension, CreationTime, LastAccessTime, LastWriteTime, FullName, Length, BaseName, Directory, PSIsContainer
                }

                $results.Add($item)
            }
        }
    }

    end {
        $results
    }
}