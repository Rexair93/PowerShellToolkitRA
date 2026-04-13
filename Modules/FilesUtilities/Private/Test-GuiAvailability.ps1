function Test-GuiAvailability {
    if (-not $IsWindows) { return $false }

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        return $false
    }
}