function ConvertTo-NormalizedExt {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline)]
        [string] $Ext
    )

    process {
        if ([string]::IsNullOrWhiteSpace($Ext)) {
            return
        }

        $Ext.Trim().TrimStart('.').ToLowerInvariant()
    }
}
