function Get-FileTypeFromBytes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [byte[]]$Bytes
    )

    function Test-Signature {
        param(
            [Parameter(Mandatory)]
            [byte[]]$InputBytes,

            [Parameter(Mandatory)]
            [byte[]]$Signature,

            [int]$Offset = 0
        )

        if ($null -eq $InputBytes -or $null -eq $Signature) { return $false }
        if ($InputBytes.Length -lt ($Offset + $Signature.Length)) { return $false }

        for ($i = 0; $i -lt $Signature.Length; $i++) {
            if ($InputBytes[$Offset + $i] -ne $Signature[$i]) {
                return $false
            }
        }

        return $true
    }

    function New-DetectionResult {
        param(
            [Parameter(Mandatory)]
            [string]$Extension,

            [Parameter(Mandatory)]
            [string]$MimeType,

            [Parameter(Mandatory)]
            [string]$Confidence,

            [Parameter(Mandatory)]
            [string]$Description
        )

        [pscustomobject]@{
            Extension   = $Extension
            MimeType    = $MimeType
            Confidence  = $Confidence
            Description = $Description
        }
    }

    if ($Bytes.Length -eq 0) {
        return New-DetectionResult -Extension 'bin' -MimeType 'application/octet-stream' -Confidence 'Low' -Description 'Byte array vuoto'
    }

    if (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x25,0x50,0x44,0x46))) {
        return New-DetectionResult -Extension 'pdf' -MimeType 'application/pdf' -Confidence 'High' -Description 'PDF'
    }

    if (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A))) {
        return New-DetectionResult -Extension 'png' -MimeType 'image/png' -Confidence 'High' -Description 'PNG'
    }

    if (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0xFF,0xD8,0xFF))) {
        return New-DetectionResult -Extension 'jpg' -MimeType 'image/jpeg' -Confidence 'High' -Description 'JPEG'
    }

    if ((Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x47,0x49,0x46,0x38,0x37,0x61))) -or (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x47,0x49,0x46,0x38,0x39,0x61)))) {
        return New-DetectionResult -Extension 'gif' -MimeType 'image/gif' -Confidence 'High' -Description 'GIF'
    }

    if (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x42,0x4D))) {
        return New-DetectionResult -Extension 'bmp' -MimeType 'image/bmp' -Confidence 'High' -Description 'BMP'
    }

    if ((Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x49,0x49,0x2A,0x00))) -or (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x4D,0x4D,0x00,0x2A)))) {
        return New-DetectionResult -Extension 'tif' -MimeType 'image/tiff' -Confidence 'High' -Description 'TIFF'
    }

    if (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x37,0x7A,0xBC,0xAF,0x27,0x1C))) {
        return New-DetectionResult -Extension '7z' -MimeType 'application/x-7z-compressed' -Confidence 'High' -Description '7-Zip'
    }

    if ((Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x52,0x61,0x72,0x21,0x1A,0x07,0x00))) -or (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x52,0x61,0x72,0x21,0x1A,0x07,0x01,0x00)))) {
        return New-DetectionResult -Extension 'rar' -MimeType 'application/vnd.rar' -Confidence 'High' -Description 'RAR'
    }

    if ((Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x50,0x4B,0x03,0x04))) -or (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x50,0x4B,0x05,0x06))) -or (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x50,0x4B,0x07,0x08)))) {

        try {
            $memoryStream = [System.IO.MemoryStream]::new($Bytes, $false)
            try {
                $zip = [System.IO.Compression.ZipArchive]::new($memoryStream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
                try {
                    $entryNames = $zip.Entries | ForEach-Object { $_.FullName }

                    if ('[Content_Types].xml' -in $entryNames) {
                        if ($entryNames -match '^word/') {
                            return New-DetectionResult -Extension 'docx' -MimeType 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' -Confidence 'High' -Description 'OOXML Word document'
                        }

                        if ($entryNames -match '^xl/') {
                            return New-DetectionResult -Extension 'xlsx' -MimeType 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' -Confidence 'High' -Description 'OOXML Excel workbook'
                        }

                        if ($entryNames -match '^ppt/') {
                            return New-DetectionResult -Extension 'pptx' -MimeType 'application/vnd.openxmlformats-officedocument.presentationml.presentation' -Confidence 'High' -Description 'OOXML PowerPoint presentation'
                        }

                        return New-DetectionResult -Extension 'zip' -MimeType 'application/zip' -Confidence 'Medium' -Description 'ZIP compatibile con struttura OOXML non distinta'
                    }

                    if ($entryNames -contains 'mimetype') {
                        return New-DetectionResult -Extension 'zip' -MimeType 'application/zip' -Confidence 'Medium' -Description 'Archivio ZIP con file mimetype'
                    }

                    return New-DetectionResult -Extension 'zip' -MimeType 'application/zip' -Confidence 'High' -Description 'Archivio ZIP'
                }
                finally {
                    $zip.Dispose()
                }
            }
            finally {
                $memoryStream.Dispose()
            }
        }
        catch {
            return New-DetectionResult -Extension 'zip' -MimeType 'application/zip' -Confidence 'Medium' -Description 'Firma ZIP rilevata ma contenuto non ispezionabile'
        }
    }

    if (Test-Signature -InputBytes $Bytes -Signature ([byte[]](0x4D,0x5A))) {
        return New-DetectionResult -Extension 'exe' -MimeType 'application/vnd.microsoft.portable-executable' -Confidence 'High' -Description 'Windows executable'
    }

    try {
        $textSampleLength = [Math]::Min($Bytes.Length, 512)
        $textSample = [System.Text.Encoding]::UTF8.GetString($Bytes, 0, $textSampleLength).TrimStart([char]0xEF,[char]0xBB,[char]0xBF).Trim()

        if ($textSample.StartsWith('{') -or $textSample.StartsWith('[')) {
            return New-DetectionResult -Extension 'json' -MimeType 'application/json' -Confidence 'Medium' -Description 'Contenuto testuale compatibile con JSON'
        }

        if ($textSample.StartsWith('<?xml') -or $textSample.StartsWith('<')) {
            return New-DetectionResult -Extension 'xml' -MimeType 'application/xml' -Confidence 'Medium' -Description 'Contenuto testuale compatibile con XML'
        }

        if (-not [string]::IsNullOrWhiteSpace($textSample)) {
            return New-DetectionResult -Extension 'txt' -MimeType 'text/plain' -Confidence 'Low' -Description 'Contenuto testuale generico'
        }
    }
    catch {
    }

    return New-DetectionResult -Extension 'bin' -MimeType 'application/octet-stream' -Confidence 'Low' -Description 'Formato non riconosciuto'
}