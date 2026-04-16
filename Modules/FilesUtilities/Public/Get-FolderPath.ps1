function Get-FolderPath {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Title = 'Seleziona una cartella',

        [Parameter()]
        [string]$InitialDirectory = (Get-Location).Path,

        [Parameter()]
        [switch]$UseConsole,

        [Parameter()]
        [switch]$AllowEmpty
    )

    if (-not $UseConsole -and (Test-GuiAvailability)) {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop

        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = $Title
        $dialog.ShowNewFolderButton = $true

        if ($InitialDirectory -and (Test-Path $InitialDirectory -PathType Container)) {
            $dialog.SelectedPath = $InitialDirectory
        }

        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $dialog.SelectedPath
        }

        if ($AllowEmpty) {
            return $null
        }

        throw 'Operazione annullata.'
    }
    else {
        Write-Information 'Modalità console attiva.'

        $prompt = if ([string]::IsNullOrWhiteSpace($InitialDirectory)) {
            $Title
        }
        else {
            "$Title [$InitialDirectory]"
        }

        $path = (Read-Host $prompt).Trim()

        if ([string]::IsNullOrWhiteSpace($path)) {
            if ($AllowEmpty) {
                return $null
            }

            if (-not [string]::IsNullOrWhiteSpace($InitialDirectory)) {
                return $InitialDirectory
            }

            throw 'Percorso non specificato.'
        }

        return $path
    }
}