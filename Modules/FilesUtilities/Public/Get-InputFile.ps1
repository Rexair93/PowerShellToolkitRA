function Get-InputFile {
    [CmdletBinding()]
    param(
        [string[]] $Formats = @("csv"),
        [string]   $Title = "Seleziona file di input",
        [string]   $InitialDirectory = (Get-Location).Path,
        [switch]   $UseConsole
    )

    $Formats = $Formats |
    ConvertTo-NormalizedExt |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
    Sort-Object -Unique

    if (-not $Formats) {
        throw "Specificare almeno un formato valido."
    }

    if (-not $UseConsole -and (Test-GuiAvailability)) {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        $dlg = [System.Windows.Forms.OpenFileDialog]::new()
        $dlg.Title = $Title
        $dlg.InitialDirectory = $InitialDirectory
        $dlg.Filter = New-FileDialogFilter -Extensions $Formats -IncludeAllFiles
        $dlg.Multiselect = $false
        $dlg.CheckFileExists = $true
        $dlg.CheckPathExists = $true
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.FileName }
        throw "Operazione annullata."
    }

    else {
        Write-Information "Modalità console attiva."
        $path = (Read-Host "Percorso file").Trim()
        if ([string]::IsNullOrWhiteSpace($path)) {
            throw "Percorso non specificato."
        }
        if (-not (Test-Path $path -PathType Leaf)) { 
            throw "File non trovato: '$path'" 
        }
        $ext = ConvertTo-NormalizedExt ([IO.Path]::GetExtension($path))
        if ($ext -notin $Formats) {
            throw "Estensione '.$ext' non consentita. Formati ammessi: $($Formats -join ', ')"
        }
        return $path
    }
}