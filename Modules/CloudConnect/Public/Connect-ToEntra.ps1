function Connect-ToEntra {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string[]] $Scopes,

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $AllowClobber,

        [Parameter()]
        [ValidateSet('Required', 'All')]
        [string] $ImportMode = 'Required'
    )

    Assert-Module -Name "Microsoft.Entra" -Scope CurrentUser -AutoInstall:$AutoInstallModules -AllowClobber:$AllowClobber

    $requiredCmdlets = @(
        'Connect-Entra',
        'Disconnect-Entra',
        'Get-EntraOrganization'
    )

    if ($ImportMode -eq 'All') {
        if (-not (Get-Module -Name Microsoft.Entra)) {
            Write-Verbose "Import completo del modulo Microsoft.Entra..."
            Import-Module Microsoft.Entra `
                -DisableNameChecking `
                -ErrorAction Stop `
                -Verbose:$false 4>$null
        }
    }
    else {
        $missingCmdlets = $requiredCmdlets | Where-Object {
            -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
        }

        if ($missingCmdlets) {
            Write-Verbose "Import selettivo del modulo Microsoft.Entra: $($requiredCmdlets -join ', ')"
            Import-Module Microsoft.Entra `
                -Cmdlet $requiredCmdlets `
                -DisableNameChecking `
                -ErrorAction Stop `
                -Verbose:$false 4>$null
        }
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

    if ($Scopes -and $Scopes.Count -gt 0) {
        Write-Verbose "Scopes specificati: $($Scopes -join ', ')"
        $connectParams.Scopes = $Scopes
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