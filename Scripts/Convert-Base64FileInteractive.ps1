<#
.SYNOPSIS
Wrapper interattivo per Convert-Base64ToFile.

.DESCRIPTION
Importa il modulo FilesUtilities e consente di selezionare un file TXT
contenente Base64, decodificandolo in un file con formato rilevato
automaticamente quando possibile.
#>
[CmdletBinding()]
param(
    [switch]$UseConsole,
    [switch]$Force
)

try {
    Import-Module FilesUtilities -ErrorAction Stop

    $result = Convert-Base64ToFile `
        -SelectInputFile `
        -UseConsole:$UseConsole `
        -Force:$Force `
        -PassThru

    Write-Host "Creato: $($result.FullName)" -ForegroundColor Green
    Write-Host "Tipo : $($result.DetectedExtension)"
    Write-Host "MIME : $($result.DetectedMimeType)"
}
catch {
    Write-Error $_
}