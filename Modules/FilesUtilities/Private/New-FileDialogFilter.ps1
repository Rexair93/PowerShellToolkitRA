function New-FileDialogFilter {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]] $Extensions,

        [Parameter()]
        [switch] $IncludeAllFiles

    )

    $friendly = @{
        "xlsx" = "Excel (*.xlsx)"
        "xls"  = "Excel 97-2003 (*.xls)"
        "csv"  = "CSV (*.csv)"
        "json" = "JSON (*.json)"
        "xml"  = "XML (*.xml)"
        "txt"  = "Testo (*.txt)"
        "html" = "HTML (*.html)"
        "pdf"  = "PDF (*.pdf)"
    }

    # Normalizza le estensioni:
    # - rimuove il punto iniziale
    # - converte in lowercase
    # - elimina duplicati e valori vuoti
    $normalized = $Extensions |
        Where-Object { $_ } |
        ForEach-Object { $_.TrimStart('.').ToLowerInvariant() } |
        Sort-Object -Unique

    
    $parts = foreach ($ext in $normalized) {
        $label = if ($friendly.ContainsKey($ext)) {
            $friendly[$ext]
        }
        else {
            "$($ext.ToUpperInvariant()) (*.$ext)"
        }

        "$label|*.$ext"
    }

    if ($IncludeAllFiles) {
        $parts += "Tutti i file (*.*)|*.*"
    }

    $parts -join "|"
}
