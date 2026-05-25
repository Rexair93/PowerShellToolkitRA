function Convert-Base64ToFile {
<#
.SYNOPSIS
Converte una stringa Base64 in un file.

.DESCRIPTION
Accetta una stringa Base64 direttamente, legge il contenuto da un file di testo
oppure consente di selezionare il file di input tramite Get-InputFile.

Se -OutputExtension non viene specificato, la funzione prova a determinare
automaticamente il formato del file analizzando i byte decodificati.

.PARAMETER Base64String
Stringa Base64 da decodificare.

.PARAMETER InputTextFile
Percorso di un file .txt contenente la stringa Base64.

.PARAMETER SelectInputFile
Apre una selezione guidata del file di input tramite Get-InputFile.

.PARAMETER OutputPath
Percorso completo del file di output.

.PARAMETER OutputExtension
Estensione da usare in output. Se omessa, la funzione prova a rilevarla
automaticamente dal contenuto.

.PARAMETER AllowedOutputExtensions
Elenco delle estensioni consentite in output.

.PARAMETER DisableAutoDetect
Disabilita il rilevamento automatico del formato e usa 'bin' se non è stata
specificata alcuna estensione esplicita.

.PARAMETER UseConsole
Forza l'uso della modalità console nei selettori file.

.PARAMETER Force
Sovrascrive il file di output se esiste già.

.PARAMETER PassThru
Restituisce informazioni dettagliate sul file creato.

.EXAMPLE
Convert-Base64ToFile -Base64String 'SGVsbG8=' -OutputPath 'C:\Temp\hello.txt'

.EXAMPLE
Convert-Base64ToFile -InputTextFile 'C:\Temp\payload.txt' -Force

.EXAMPLE
Convert-Base64ToFile -SelectInputFile -UseConsole -PassThru

.EXAMPLE
Convert-Base64ToFile -InputTextFile 'C:\Temp\file64.txt' -DisableAutoDetect -OutputPath 'C:\Temp\output.bin'
#>
    [CmdletBinding(
        SupportsShouldProcess,
        DefaultParameterSetName = 'FromString',
        ConfirmImpact = 'Medium'
    )]
    param(
        [Parameter(
            Mandatory,
            ParameterSetName = 'FromString',
            Position = 0,
            ValueFromPipeline,
            ValueFromPipelineByPropertyName
        )]
        [ValidateNotNullOrEmpty()]
        [string]$Base64String,

        [Parameter(
            Mandatory,
            ParameterSetName = 'FromFile',
            Position = 0
        )]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({
            if (-not (Test-Path -Path $_ -PathType Leaf)) {
                throw "File non trovato: $_"
            }

            $ext = [IO.Path]::GetExtension($_).TrimStart('.').ToLowerInvariant()
            if ($ext -ne 'txt') {
                throw "Il file di input deve avere estensione .txt: $_"
            }

            $true
        })]
        [string]$InputTextFile,

        [Parameter(
            Mandatory,
            ParameterSetName = 'FromPicker'
        )]
        [switch]$SelectInputFile,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$OutputPath,

        [Parameter()]
        [ValidatePattern('^[.]?[a-zA-Z0-9]{1,10}$')]
        [string]$OutputExtension,

        [Parameter()]
        [AllowNull()]
        [ValidateNotNullOrEmpty()]
        [string[]]$AllowedOutputExtensions = @(
            'bin','txt','json','xml',
            'pdf','png','jpg','jpeg','gif','bmp','tif','tiff',
            'zip','docx','xlsx','pptx','7z','rar','exe'
        ),

        [Parameter()]
        [switch]$DisableAutoDetect,

        [Parameter()]
        [switch]$UseConsole,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$PassThru
    )

    begin {
        $knownFormats = @(
            'bin','txt','json','xml',
            'pdf','png','jpg','jpeg','gif','bmp','tif','tiff',
            'zip','docx','xlsx','pptx','7z','rar','exe'
        )

        function Get-NormalizedExtension {
            param([AllowNull()][AllowEmptyString()][string]$Extension)

            if ([string]::IsNullOrWhiteSpace($Extension)) {
                return $null
            }

            return $Extension.Trim().TrimStart('.').ToLowerInvariant()
        }

        function Get-NormalizedBase64Text {
            param([Parameter(Mandatory)][string]$Text)

            $value = $Text.Trim()

            if ($value -match '^data:.*?;base64,') {
                $value = $value -replace '^data:.*?;base64,', ''
            }

            $value = $value -replace '\s', ''
            return $value
        }

        function Read-Base64FromTextFile {
            param([Parameter(Mandatory)][string]$Path)

            $content = Get-Content -Path $Path -Raw -Encoding UTF8
            if ([string]::IsNullOrWhiteSpace($content)) {
                throw "Il file di input è vuoto: $Path"
            }

            return (Get-NormalizedBase64Text -Text $content)
        }
    }

    process {
        $hasAllowedFilter = $PSBoundParameters.ContainsKey('AllowedOutputExtensions')

        $knownFormatsNormalized = $knownFormats |
            ForEach-Object { Get-NormalizedExtension $_ } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique

        $allowedFormats = @()
        if ($hasAllowedFilter) {
            $allowedFormats = $AllowedOutputExtensions |
                ForEach-Object { Get-NormalizedExtension $_ } |
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                Sort-Object -Unique

        if (-not $allowedFormats -or $allowedFormats.Count -eq 0) {
                throw "AllowedOutputExtensions è stato specificato ma non contiene formati validi."
            }
        }

        $normalizedOutputExtension = Get-NormalizedExtension $OutputExtension
        if ($normalizedOutputExtension -and $hasAllowedFilter -and $normalizedOutputExtension -notin $allowedFormats) {
            throw "L'estensione '$normalizedOutputExtension' non è ammessa. Formati consentiti: $($allowedFormats -join ', ')"
        }

        switch ($PSCmdlet.ParameterSetName) {
            'FromString' {
                $normalizedBase64 = Get-NormalizedBase64Text -Text $Base64String
                $inputSource = 'Base64String'
            }

            'FromFile' {
                $normalizedBase64 = Read-Base64FromTextFile -Path $InputTextFile
                $inputSource = $InputTextFile
            }

            'FromPicker' {
                if (-not (Get-Command Get-InputFile -ErrorAction SilentlyContinue)) {
                    throw "La funzione Get-InputFile non è disponibile nel modulo corrente."
                }

                $selectedFile = Get-InputFile `
                    -Formats @('txt') `
                    -Title 'Seleziona il file TXT contenente il Base64' `
                    -InitialDirectory (Get-Location).Path `
                    -UseConsole:$UseConsole

                $normalizedBase64 = Read-Base64FromTextFile -Path $selectedFile
                $inputSource = $selectedFile
            }

            default {
                throw "ParameterSet non gestito: $($PSCmdlet.ParameterSetName)"
            }
        }

        if ([string]::IsNullOrWhiteSpace($normalizedBase64)) {
            throw "La stringa Base64 è vuota dopo la normalizzazione."
        }

        try {
            $bytes = [Convert]::FromBase64String($normalizedBase64)
        }
        catch {
            throw "La stringa fornita non è un Base64 valido. $($_.Exception.Message)"
        }

        if ($bytes.Length -eq 0) {
            throw "La decodifica ha prodotto un contenuto vuoto."
        }

        $detectedType = $null
        if (-not $normalizedOutputExtension -and -not $DisableAutoDetect) {
            if (-not (Get-Command Get-FileTypeFromBytes -ErrorAction SilentlyContinue)) {
                throw "La funzione privata Get-FileTypeFromBytes non è disponibile nel modulo corrente."
            }

            $detectedType = Get-FileTypeFromBytes -Bytes $bytes
            $normalizedOutputExtension = Get-NormalizedExtension $detectedType.Extension

            if ([string]::IsNullOrWhiteSpace($normalizedOutputExtension)) {
                $normalizedOutputExtension = 'bin'
            }

            Write-Verbose ("Formato rilevato: {0} | MIME: {1} | Confidenza: {2}" -f
                $detectedType.Extension,
                $detectedType.MimeType,
                $detectedType.Confidence)
        }

        if (-not $normalizedOutputExtension) {
            $normalizedOutputExtension = 'bin'
        }

        if ($hasAllowedFilter -and $normalizedOutputExtension -notin $allowedFormats) {
            throw "L'estensione finale '$normalizedOutputExtension' non è inclusa in AllowedOutputExtensions."
        }

        if (-not $OutputPath) {
            $defaultBaseName = 'decoded-output'
            if (Get-Command ConvertTo-SafeFileName -ErrorAction SilentlyContinue) {
                $defaultBaseName = ConvertTo-SafeFileName -Name $defaultBaseName
            }

            $defaultFileName = '{0}.{1}' -f $defaultBaseName, $normalizedOutputExtension

            $dialogFormats = if ($hasAllowedFilter) { $allowedFormats } else { $knownFormatsNormalized }

            if (Get-Command Get-ExportDestination -ErrorAction SilentlyContinue) {
                $OutputPath = Get-ExportDestination `
                    -DefaultFileName $defaultFileName `
                    -InitialDirectory (Get-Location).Path `
                    -Formats $dialogFormats `
                    -PreferredFormat $normalizedOutputExtension `
                    -Title 'Scegli il file di output' `
                    -UseConsole:$UseConsole `
                    -Force:$Force `
                    -AllowAnyExtension: $true `
                    -AsString
            }
            else {
                $OutputPath = Join-Path -Path (Get-Location).Path -ChildPath $defaultFileName
            }
        }

        $existingExt = Get-NormalizedExtension ([IO.Path]::GetExtension($OutputPath))
        if (-not $existingExt) {
            $OutputPath = '{0}.{1}' -f $OutputPath, $normalizedOutputExtension
            $existingExt = Get-NormalizedExtension ([IO.Path]::GetExtension($OutputPath))
        }

        if ($hasAllowedFilter -and$existingExt -notin $allowedFormats) {
            throw "Estensione finale '.$existingExt' non consentita. Ammesse: $($allowedFormats -join ', ')"
        }

        if ($detectedType -and $existingExt -and $existingExt -ne $detectedType.Extension -and -not $hasAllowedFilter) {
            Write-Warning "L'estensione scelta '.$existingExt' differisce dal formato rilevato '$($detectedType.Extension)'."
        }

        $directory = [IO.Path]::GetDirectoryName($OutputPath)
        if ($directory -and -not (Test-Path -Path $directory -PathType Container)) {
            New-Item -Path $directory -ItemType Directory -Force | Out-Null
        }

        if ((Test-Path -Path $OutputPath -PathType Leaf) -and -not $Force) {
            throw "Il file di output esiste già: $OutputPath. Usare -Force per sovrascrivere."
        }

        if ($PSCmdlet.ShouldProcess($OutputPath, 'Scrittura file decodificato da Base64')) {
            [IO.File]::WriteAllBytes($OutputPath, $bytes)
        }

        $item = Get-Item -Path $OutputPath -ErrorAction Stop

        if ($PassThru) {
            [pscustomobject]@{
                FullName            = $item.FullName
                Name                = $item.Name
                Length              = $item.Length
                Extension           = $item.Extension
                LastWriteTime       = $item.LastWriteTime
                InputSource         = $inputSource
                DetectedExtension   = if ($detectedType) { $detectedType.Extension } else { $null }
                DetectedMimeType    = if ($detectedType) { $detectedType.MimeType } else { $null }
                DetectionConfidence = if ($detectedType) { $detectedType.Confidence } else { $null }
            }
        }
        else {
            $item
        }
    }
}