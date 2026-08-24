function Get-FileMediaTagMap {
    <#
    .SYNOPSIS
        Elenca i metadati disponibili per un file tramite Shell.Application ed ExifTool.

    .DESCRIPTION
        Legge le proprietà esposte da Windows Shell e, se disponibile, i tag embedded
        restituiti da ExifTool. È utile per identificare i nomi da usare con
        Get-FolderContentReport e con il parametro CustomMediaTags.

    .PARAMETER FilePath
        Percorso del file da analizzare.

    .PARAMETER ExifToolPath
        Percorso esplicito di exiftool.exe. Se omesso, viene cercato nel PATH.

    .PARAMETER IncludeEmpty
        Include anche le proprietà Shell e i tag ExifTool privi di valore.

    .PARAMETER Source
        Limita la ricerca a Shell, ExifTool oppure a entrambi.

    .EXAMPLE
        Get-FileMediaTagMap -FilePath 'D:\Media\brano01.mp3'

    .EXAMPLE
        Get-FileMediaTagMap -FilePath 'D:\Media\brano01.mp3' -Source ExifTool

    .EXAMPLE
        Get-FileMediaTagMap -FilePath 'D:\Media\brano01.mp3' |
            Where-Object { $_.Source -eq 'ExifTool' } |
            Format-Table -AutoSize

    .NOTES
        Shell.Application è disponibile su Windows.
        ExifTool deve essere installato e disponibile nel PATH oppure indicato tramite ExifToolPath.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('FullName', 'Path')]
        [string] $FilePath,

        [string] $ExifToolPath = 'C:\exiftool\exiftool.exe',

        [switch] $IncludeEmpty,

        [ValidateSet('Shell', 'ExifTool', 'All')]
        [string] $Source = 'All'
    )

    begin {
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

        function Get-ShellMetadata {
            param(
                [Parameter(Mandatory)]
                [string] $ResolvedFilePath,

                [switch] $ReturnEmpty
            )

            $shell = $null

            try {
                $parentPath = Split-Path -Path $ResolvedFilePath -Parent
                $fileName = Split-Path -Path $ResolvedFilePath -Leaf

                $shell = New-Object -ComObject Shell.Application
                $folder = $shell.Namespace($parentPath)
                $item = $folder.ParseName($fileName)

                if ($null -eq $folder -or $null -eq $item) {
                    Write-Warning "Impossibile inizializzare la Shell per '$ResolvedFilePath'."
                    return
                }

                for ($index = 0; $index -le 400; $index++) {
                    $propertyName = $folder.GetDetailsOf($null, $index)
                    $propertyValue = $folder.GetDetailsOf($item, $index)

                    if ([string]::IsNullOrWhiteSpace($propertyName)) {
                        continue
                    }

                    if (-not $ReturnEmpty -and [string]::IsNullOrWhiteSpace($propertyValue)) {
                        continue
                    }

                    [pscustomobject]@{
                        Source              = 'Shell'
                        Index               = $index
                        Property            = $propertyName
                        ExifToolName        = $null
                        Value               = if ([string]::IsNullOrWhiteSpace($propertyValue)) {
                            $null
                        }
                        else {
                            $propertyValue.Trim()
                        }
                        SuggestedCustomTag = $propertyName
                    }
                }
            }
            catch {
                Write-Warning "Errore nella lettura delle proprietà Shell per '$ResolvedFilePath': $($_.Exception.Message)"
            }
            finally {
                if ($null -ne $shell) {
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
                }
            }
        }

        function Get-ExifToolMetadata {
            param(
                [Parameter(Mandatory)]
                [string] $ResolvedFilePath,

                [Parameter(Mandatory)]
                [string] $ToolPath,

                [switch] $ReturnEmpty
            )

            try {
                $arguments = @(
                    '-j',
                    '-G1',
                    '-a',
                    '-s',
                    '-duplicates',
                    $ResolvedFilePath
                )

                $jsonOutput = & $ToolPath @arguments 2>$null
                $jsonText = $jsonOutput -join [Environment]::NewLine

                if ([string]::IsNullOrWhiteSpace($jsonText)) {
                    Write-Warning "ExifTool non ha restituito dati per '$ResolvedFilePath'."
                    return
                }

                $parsedOutput = $jsonText | ConvertFrom-Json

                if ($parsedOutput -is [array]) {
                    $parsedOutput = $parsedOutput | Select-Object -First 1
                }

                foreach ($property in $parsedOutput.PSObject.Properties) {
                    $propertyValue = $property.Value

                    if (-not $ReturnEmpty) {
                        if ($null -eq $propertyValue) {
                            continue
                        }

                        if ($propertyValue -is [array]) {
                            if ($propertyValue.Count -eq 0) {
                                continue
                            }
                        }
                        elseif ([string]::IsNullOrWhiteSpace([string] $propertyValue)) {
                            continue
                        }
                    }

                    $displayValue = if ($propertyValue -is [array]) {
                        $propertyValue -join '; '
                    }
                    else {
                        [string] $propertyValue
                    }

                    [pscustomobject]@{
                        Source              = 'ExifTool'
                        Index               = $null
                        Property            = ($property.Name -replace '^.*:', '')
                        ExifToolName        = $property.Name
                        Value               = $displayValue
                        SuggestedCustomTag = ($property.Name -replace '^.*:', '')
                    }
                }
            }
            catch {
                Write-Warning "Errore nella lettura ExifTool per '$ResolvedFilePath': $($_.Exception.Message)"
            }
        }
    }

    process {
        if (-not (Test-Path -Path $FilePath -PathType Leaf)) {
            Write-Warning "Il file '$FilePath' non esiste o non è accessibile."
            return
        }

        $resolvedFilePath = (Resolve-Path -Path $FilePath).Path

        if ($Source -in @('Shell', 'All')) {
            Get-ShellMetadata `
                -ResolvedFilePath $resolvedFilePath `
                -ReturnEmpty:$IncludeEmpty
        }

        if ($Source -in @('ExifTool', 'All')) {
            $effectiveExifToolPath = Resolve-ExifToolExecutable -PreferredPath $ExifToolPath

            if ($null -eq $effectiveExifToolPath) {
                Write-Verbose "ExifTool non trovato. Per abilitarlo installalo o specifica -ExifToolPath."
            }
            else {
                Get-ExifToolMetadata `
                    -ResolvedFilePath $resolvedFilePath `
                    -ToolPath $effectiveExifToolPath `
                    -ReturnEmpty:$IncludeEmpty
            }
        }
    }
}