function Get-TeamsAllowedDomains {
    <#
    .SYNOPSIS
        Recupera i domini consentiti della federazione Microsoft Teams.

    .DESCRIPTION
        Si connette a Microsoft Teams tramite Connect-ToMicrosoftTeams e restituisce
        i domini consentiti dalla configurazione della federazione come oggetti PowerShell.

    .PARAMETER TenantId
        Tenant ID da usare per la connessione a Microsoft Teams.

    .PARAMETER UseDeviceCode
        Prova a usare l'autenticazione device code, se supportata dal modulo MicrosoftTeams.

    .PARAMETER ForceReconnect
        Forza la disconnessione e una nuova connessione a Microsoft Teams.

    .PARAMETER AutoInstallModules
        Installa automaticamente i moduli mancanti richiesti.

    .OUTPUTS
        PSCustomObject con proprietà Domain per ogni dominio consentito trovato.

    .EXAMPLE
        Get-TeamsAllowedDomains

    .EXAMPLE
        Get-TeamsAllowedDomains -TenantId "contoso.onmicrosoft.com" -UseDeviceCode -Verbose
    #>
    [CmdletBinding()]
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

    Write-Verbose "Connessione a Microsoft Teams..."
    Connect-ToMicrosoftTeams `
        -TenantId           $TenantId `
        -UseDeviceCode:     $UseDeviceCode `
        -ForceReconnect:    $ForceReconnect `
        -AutoInstallModules:$AutoInstallModules `
        -Verbose:           $VerbosePreference

    Write-Verbose "Recupero configurazione domini consentiti da Teams..."
    $domains = Get-CsTenantFederationConfiguration -ErrorAction Stop |
        Select-Object -ExpandProperty AllowedDomains -ErrorAction Stop |
        Select-Object -ExpandProperty AllowedDomain  -ErrorAction Stop |
        Select-Object -ExpandProperty Domain |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Sort-Object -Unique |
        ForEach-Object {
            [pscustomobject]@{ Domain = $_ }
        }

    if (-not $domains) {
        Write-Warning "Nessun dominio consentito trovato nella configurazione Teams."
        return
    }

    $domains
}