function Connect-ToMicrosoftTeams {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [switch] $AutoInstallModules
    )

    Assert-Module -Name "MicrosoftTeams" -Scope CurrentUser -AutoInstall:$AutoInstallModules

    # In alcune versioni del modulo Teams c'è già una sessione implicita.
    # Se non vuoi riconnettere, prova un comando leggero.
    if (-not $ForceReconnect) {
        try {
            Write-Verbose "Verifica sessione Microsoft Teams esistente..."
            Get-CsTenant -ErrorAction Stop | Out-Null
            Write-Verbose "Sessione Microsoft Teams già attiva."
            return
        } catch {
            Write-Verbose "Nessuna sessione valida trovata."
         }
    }

    
    if ($ForceReconnect) {
        Write-Verbose "Disconnessione forzata da Microsoft Teams."
        Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
    }

    $connectParams = @{}
    if ($TenantId) { $connectParams.TenantId = $TenantId }

    # UseDeviceCode potrebbe non esistere in alcune versioni del modulo
    if ($UseDeviceCode) {
        
        $connectCmd = Get-Command Connect-MicrosoftTeams -ErrorAction Stop

        if ($connectCmd.Parameters.ContainsKey('UseDeviceAuthentication')) {
            Write-Verbose "Uso -UseDeviceAuthentication."
            $connectParams.UseDeviceAuthentication = $true
        }
        elseif ($connectCmd.Parameters.ContainsKey('UseDeviceCode')) {
            Write-Verbose "Uso -UseDeviceCode."
            $connectParams.UseDeviceCode = $true
        }
        else {
            Write-Warning "Il modulo MicrosoftTeams installato non supporta il device code. Procedo senza."
        }
    }

    if ($PSCmdlet.ShouldProcess("Microsoft Teams", "Connect")) {
        Write-Verbose "Connessione a Microsoft Teams..."
        Connect-MicrosoftTeams @connectParams | Out-Null
    }
}
