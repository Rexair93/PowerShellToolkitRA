function Connect-ToGraph {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string[]] $Scopes,

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [switch] $AutoInstallModules
    )

    Assert-Module -Name "Microsoft.Graph.Authentication" -Scope CurrentUser -AutoInstall:$AutoInstallModules

    if (-not (Get-Module -Name Microsoft.Graph.Authentication)) {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    }

    # Se già connesso e non vuoi riconnettere, esci
    
    $ctx = $null
    try {
        $ctx = Get-MgContext -ErrorAction Stop
    } catch {}

    if ($ctx -and -not $ForceReconnect) {
        $missingScopes = $Scopes | Where-Object { $_ -notin $ctx.Scopes }
        if (-not $missingScopes) {
            Write-Verbose "Connessione a Microsoft Graph già valida."
            return
        }

        Write-Verbose "Scope mancanti: $($missingScopes -join ', ')"
    }

    if ($ForceReconnect -and $ctx) {
        Write-Verbose "Disconnessione da Microsoft Graph."
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    }

    $connectParams = @{
        Scopes    = $Scopes
        NoWelcome = $true
    }
    if ($TenantId)      { $connectParams.TenantId = $TenantId }
    if ($UseDeviceCode) { $connectParams.UseDeviceCode = $true }

    if ($PSCmdlet.ShouldProcess("Microsoft Graph", "Connect")) {
        Write-Verbose "Connessione a Microsoft Graph..."
        Connect-MgGraph @connectParams | Out-Null
    }
}
