function Assert-Module {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [string] $Name,

        [Parameter()]
        [string] $MinimumVersion,

        [Parameter()]
        [ValidateSet('CurrentUser', 'AllUsers')]
        [string] $Scope = 'CurrentUser',

        [Parameter()]
        [switch] $AutoInstall,

        [Parameter()]
        [switch] $AllowClobber
    )

    # Trova la versione più recente installata
    $installed = Get-Module -ListAvailable -Name $Name |
    Sort-Object Version -Descending |
    Select-Object -First 1

    
    $needsInstall = $false

    if (-not $installed) {
        $needsInstall = $true
    }
    elseif ($MinimumVersion -and $installed.Version -lt $MinimumVersion) {
        $needsInstall = $true
    }

    if (-not $needsInstall) {
        return
    }

    if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
        throw "Modulo '$Name' mancante e Install-Module non disponibile. Aggiorna PowerShellGet."
    }

    $installParams = @{
        Name        = $Name
        Scope       = $Scope
        ErrorAction = 'Stop'
        Force       = $true
    }

    if ($MinimumVersion) {
        $installParams.MinimumVersion = $MinimumVersion
    }

    if ($AllowClobber) {
        $installParams.AllowClobber = $true
    }

    if ($AutoInstall) {
        Install-Module @installParams
        return
    }

    $question = if ($MinimumVersion) {
        "Installare il modulo '$Name' (AllowClobber=$AllowClobber) (>= $MinimumVersion)?"
    }
    else {
        "Installare il modulo '$Name' (AllowClobber=$AllowClobber)?"
    }

    if ($PSCmdlet.ShouldContinue($question, "Modulo mancante")) {
        Install-Module @installParams
    }
    else {
        throw "Modulo '$Name' non installato per scelta dell'utente."
    }
}