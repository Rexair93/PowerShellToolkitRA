function Connect-ToEntra {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $AllowClobber
    )

    Assert-Module -Name "Microsoft.Entra" -Scope CurrentUser -AutoInstall:$AutoInstallModules -AllowClobber:$AllowClobber

    if (-not (Get-Module -Name Microsoft.Entra)) {
        Import-Module Microsoft.Entra -ErrorAction Stop
    }

    # -----------------------------------------------------------------
    # Verifica connessione esistente (non esiste un vero "context")
    if (-not $ForceReconnect) {
        try {
            Write-Verbose "Verifica connessione Microsoft Entra esistente..."
            Get-EntraOrganization -ErrorAction Stop | Out-Null
            Write-Verbose "Connessione Microsoft Entra già attiva."
            return
        }
        catch {
            Write-Verbose "Nessuna connessione Entra valida rilevata."
        }
    }

    # -----------------------------------------------------------------
    # Disconnessione esplicita (best effort)
    if ($ForceReconnect) {
        Write-Verbose "Disconnessione forzata da Microsoft Entra."
        Disconnect-Entra -ErrorAction SilentlyContinue
    }

    # -----------------------------------------------------------------
    # Costruzione parametri di connessione
    $connectParams = @{}

    if ($TenantId) {
        $connectParams.TenantId = $TenantId
    }

    if ($UseDeviceCode) {
        $cmd = Get-Command Connect-Entra -ErrorAction Stop

        if ($cmd.Parameters.ContainsKey('UseDeviceAuthentication')) {
            Write-Verbose "Uso autenticazione tramite device code."
            $connectParams.UseDeviceAuthentication = $true
        }
        else {
            Write-Warning "La versione del modulo Microsoft.Entra non supporta il device code."
        }
    }

    # -----------------------------------------------------------------
    if ($PSCmdlet.ShouldProcess("Microsoft Entra", "Connect")) {
        Write-Verbose "Connessione a Microsoft Entra..."
        Connect-Entra @connectParams | Out-Null
    }
}