function Get-SubFolderList {
    <#
    .SYNOPSIS
        Restituisce l'elenco delle sottocartelle di primo livello di una cartella specificata.

    .DESCRIPTION
        Dato un percorso di cartella, restituisce un array di oggetti con il nome
        e il percorso completo di ogni sottocartella di primo livello (non ricorsivo).

    .PARAMETER FolderPath
        Percorso della cartella principale da analizzare.

    .OUTPUTS
        Array di [pscustomobject] con proprietà: Name, FullName.

    .EXAMPLE
        Get-SubFolderList -FolderPath "C:\Progetti"
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $FolderPath
    )

    if (-not (Test-Path -Path $FolderPath -PathType Container)) {
        throw "Il percorso '$FolderPath' non esiste o non è una cartella."
    }

    $resolved = (Resolve-Path -Path $FolderPath).Path

    Get-ChildItem -Path $resolved -Directory -Force |
        Select-Object Name, FullName
}