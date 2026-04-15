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
        [switch] $AutoInstallModules,

        [Parameter()]
        [ValidateSet('Required', 'All')]
        [string] $ImportMode = 'Required'
    )

    Assert-Module -Name "Microsoft.Graph.Authentication" -Scope CurrentUser -AutoInstall:$AutoInstallModules

    $requiredCmdlets = @(
        'Connect-MgGraph',
        'Disconnect-MgGraph',
        'Get-MgContext'
    )

    if ($ImportMode -eq 'All') {
        if (-not (Get-Module -Name Microsoft.Graph.Authentication)) {
            Write-Verbose "Import completo del modulo Microsoft.Graph.Authentication..."
            Import-Module Microsoft.Graph.Authentication `
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
            Write-Verbose "Import selettivo del modulo Microsoft.Graph.Authentication: $($requiredCmdlets -join ', ')"
            Import-Module Microsoft.Graph.Authentication `
                -Cmdlet $requiredCmdlets `
                -DisableNameChecking `
                -ErrorAction Stop `
                -Verbose:$false 4>$null
        }
    }
    # Se già connesso e non vuoi riconnettere, esci
    
    $ctx = $null
    try {
        $ctx = Get-MgContext -ErrorAction Stop
    }
    catch {
        $ctx = $null
    }

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
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    if ($UseDeviceCode) { $connectParams.UseDeviceCode = $true }

    if ($PSCmdlet.ShouldProcess("Microsoft Graph", "Connect")) {
        Write-Verbose "Connessione a Microsoft Graph..."
        Connect-MgGraph @connectParams | Out-Null
    }
}
