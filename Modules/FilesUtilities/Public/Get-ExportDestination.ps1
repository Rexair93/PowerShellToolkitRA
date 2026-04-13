function Get-ExportDestination {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [string] $DefaultFileName = "report.xlsx",
        [string] $InitialDirectory = (Get-Location).Path,
        [ValidateNotNullOrEmpty()]
        [string[]] $Formats = @("xlsx","csv"),
        [string] $PreferredFormat,
        [string] $Title = "Scegli dove salvare il report",
        [switch] $UseConsole,
        [switch] $Force,
        [switch] $AsString
    )

    # --- Normalizzazione formati ---
    $Formats = $Formats |
        ConvertTo-NormalizedExt |
        Where-Object { $_ } |
        Sort-Object -Unique

    if (-not $Formats) {
        throw "Specificare almeno un formato valido."
    }

    $PreferredFormat = ConvertTo-NormalizedExt $PreferredFormat
    if ($PreferredFormat -and ($PreferredFormat -notin $Formats)) {
        throw "PreferredFormat '$PreferredFormat' non valido."
    }

    # --- Costruzione default file ---
    $defaultNameExt = ConvertTo-NormalizedExt ([IO.Path]::GetExtension($DefaultFileName))
    $defaultBase    = if ($defaultNameExt) {
        [IO.Path]::GetFileNameWithoutExtension($DefaultFileName)
    } else {
        $DefaultFileName
    }

    $defaultExt  = $PreferredFormat ?? ($defaultNameExt -in $Formats ? $defaultNameExt : $Formats[0])
    $defaultFile = "$defaultBase.$defaultExt"
    $defaultFull = Join-Path $InitialDirectory $defaultFile

    # --- GUI ---
    if (-not $UseConsole -and (Test-GuiAvailability)) {
        $dlg = [System.Windows.Forms.SaveFileDialog]::new()
        $dlg.Title            = $Title
        $dlg.InitialDirectory = $InitialDirectory
        $dlg.Filter           = New-FileDialogFilter -Extensions $Formats -IncludeAllFiles
        $dlg.FileName         = $defaultFile
        $dlg.DefaultExt       = $defaultExt
        $dlg.AddExtension     = $true
        $dlg.OverwritePrompt = $true

        if ($dlg.ShowDialog() -eq 'OK') {
            $path = $dlg.FileName
        } else {
            throw "Operazione annullata."
        }
    }
    else {
        Write-Information "Modalità console attiva."
        $path = Read-Host "Percorso file [$defaultFull]"
        if (-not $path) { $path = $defaultFull }
    }

    # --- Normalizzazione path finale ---
    if (Test-Path $path -PathType Container) {
        $path = Join-Path $path $defaultFile
    }

    $ext = ConvertTo-NormalizedExt ([IO.Path]::GetExtension($path))
    if (-not $ext) {
        $ext = $defaultExt
        $path = "$path.$ext"
    }

    if ($ext -notin $Formats) {
        throw "Estensione '.$ext' non consentita."
    }

    if ($PSCmdlet.ShouldProcess($path, "Creazione file export")) {

        $dir = [IO.Path]::GetDirectoryName($path)
        if ($dir -and -not (Test-Path $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        if ((Test-Path $path) -and -not $Force -and $UseConsole) {
            if ((Read-Host "Sovrascrivere? (S/N)") -notmatch '^[SsYy]') {
                throw "Operazione annullata."
            }
        }
    }

    $out = [pscustomobject]@{ Path = $path; Format = $ext }
    return $AsString ? $out.Path : $out
}