<#
.SYNOPSIS
    Esporta i domini consentiti della federazione Microsoft Teams in un file XLSX o CSV.

.DESCRIPTION
    Recupera i domini tramite Get-TeamsAllowedDomains e li salva nel percorso
    specificato. Se OutputPath è omesso, viene richiesto tramite Get-ExportDestination
    (GUI o console).

.PARAMETER OutputPath
    Percorso completo del file di output. Se omesso, viene richiesto interattivamente.

.PARAMETER TenantId
    Tenant ID da usare per la connessione a Microsoft Teams.

.PARAMETER UseDeviceCode
    Usa l'autenticazione device code se supportata.

.PARAMETER UseConsole
    Forza la selezione del percorso in modalità console.

.PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti.

.PARAMETER ForceReconnect
    Forza una nuova connessione a Microsoft Teams.

.PARAMETER Force
    Sovrascrive il file di output se già esistente.

.EXAMPLE
    .\Export-TeamsAllowedDomains.ps1

.EXAMPLE
    .\Export-TeamsAllowedDomains.ps1 -OutputPath .\domini.xlsx -UseConsole -Verbose
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [switch] $UseDeviceCode,

    [Parameter()]
    [switch] $UseConsole,

    [Parameter()]
    [switch] $AutoInstallModules,

    [Parameter()]
    [switch] $ForceReconnect,

    [Parameter()]
    [switch] $Force
)

Import-Module CloudOperations -ErrorAction Stop
Import-Module FilesUtilities  -ErrorAction Stop

try {
    # -----------------------------------------------------------------------
    # 1. Recupera i dati da Teams
    # -----------------------------------------------------------------------
    $results = Get-TeamsAllowedDomains `
        -TenantId:          $TenantId `
        -UseDeviceCode:     $UseDeviceCode `
        -ForceReconnect:    $ForceReconnect `
        -AutoInstallModules:$AutoInstallModules `
        -Verbose:           $VerbosePreference

    if (-not $results) {
        exit 0
    }

    # -----------------------------------------------------------------------
    # 2. Risolvi il percorso di output
    # -----------------------------------------------------------------------
    if (-not $OutputPath) {
        Write-Verbose "Richiesta percorso di output..."
        $destination = Get-ExportDestination `
            -DefaultFileName "Teams-AllowedDomains.xlsx" `
            -Formats         @('xlsx', 'csv') `
            -PreferredFormat 'xlsx' `
            -Title           "Scegli dove salvare i domini consentiti di Teams" `
            -UseConsole:     $UseConsole `
            -Force:          $Force
        $OutputPath = $destination.Path
    }

    # -----------------------------------------------------------------------
    # 3. Esporta
    # -----------------------------------------------------------------------
    if ($PSCmdlet.ShouldProcess($OutputPath, "Esporta domini consentiti Teams")) {
        Write-Verbose "Esportazione in corso verso '$OutputPath'..."
        $export = Export-Results `
            -InputObject   $results `
            -Path          $OutputPath `
            -WorksheetName 'AllowedDomains' `
            -Force:        $Force

        Write-Host "Esportazione completata: $($export.Path) [$($export.Format.ToUpper())]" -ForegroundColor Green
    }
}
catch {
    if ($_.Exception.Message -eq 'Operazione annullata.') {
        Write-Warning "Operazione annullata dall'utente."
        exit 0
    }

    Write-Error $_.Exception.Message
    exit 1
}