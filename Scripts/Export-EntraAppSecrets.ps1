<#
.SYNOPSIS
Esporta i secret delle app registrate in Entra ID il cui nome contiene una stringa.

.DESCRIPTION
Importa automaticamente i moduli custom necessari, richiama la funzione
Get-EntraAppSecrets per cercare le application registrations il cui
DisplayName contiene una stringa specificata e ne esporta i risultati
tramite Export-Results.

Le colonne esportate sono:
- Display Name
- Application (Client) ID
- Data Scadenza Secret
- Secret ID

Se SearchString non viene specificato, lo script chiede all'utente di inserirlo.
Se OutputPath non viene specificato, il percorso viene richiesto tramite
Get-ExportDestination.

.PARAMETER SearchString
Stringa da cercare nel DisplayName delle app registrate.

.PARAMETER OutputPath
Percorso completo del file di output. Se omesso, viene richiesto tramite
Get-ExportDestination.

.PARAMETER TenantId
Tenant ID da usare per la connessione a Microsoft Graph.

.PARAMETER UseDeviceCode
Usa l'autenticazione device code per la connessione a Microsoft Graph.

.PARAMETER UseConsole
Forza la selezione dei percorsi in modalità console.

.PARAMETER AutoInstallModules
Installa automaticamente i moduli mancanti richiesti.

.PARAMETER ForceReconnect
Forza una nuova connessione a Microsoft Graph.

.PARAMETER Force
Consente la sovrascrittura del file di output quando supportato.

.PARAMETER PageSize
Numero di elementi per pagina richiesti a Graph.

.PARAMETER MaxResults
Numero massimo di application da elaborare dopo la ricerca Graph.
0 = nessun limite.

.EXAMPLE
.\Export-EntraAppSecrets.ps1 -SearchString "Contoso"

.EXAMPLE
.\Export-EntraAppSecrets.ps1

.EXAMPLE
.\Export-EntraAppSecrets.ps1 -SearchString "API" -Verbose
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [AllowEmptyString()]
    [string] $SearchString,

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
    [switch] $Force,

    [Parameter()]
    [ValidateRange(1, 999)]
    [int] $PageSize = 100,

    [Parameter()]
    [ValidateRange(0, 1000000)]
    [int] $MaxResults = 1000
)

$scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path
$toolkitRoot = Split-Path -Parent $scriptDirectory

$moduleCandidates = @{
    CloudConnect   = @(
        (Join-Path $toolkitRoot 'CloudConnect\CloudConnect.psm1'),
        (Join-Path $toolkitRoot 'Modules\CloudConnect\CloudConnect.psm1')
    )
    CloudOperations = @(
        (Join-Path $toolkitRoot 'CloudOperations\CloudOperations.psm1'),
        (Join-Path $toolkitRoot 'Modules\CloudOperations\CloudOperations.psm1')
    )
    FilesUtilities = @(
        (Join-Path $toolkitRoot 'FilesUtilities\FilesUtilities.psm1'),
        (Join-Path $toolkitRoot 'Modules\FilesUtilities\FilesUtilities.psm1')
    )
}

foreach ($moduleName in $moduleCandidates.Keys) {
    if (-not (Get-Module -Name $moduleName)) {
        $modulePath = $moduleCandidates[$moduleName] |
            Where-Object { Test-Path $_ } |
            Select-Object -First 1

        if (-not $modulePath) {
            throw "Modulo '$moduleName' non trovato. Percorsi provati: $($moduleCandidates[$moduleName] -join '; ')"
        }

        Write-Verbose "Import modulo $moduleName da '$modulePath'..."
        Import-Module $modulePath -Force -ErrorAction Stop
    }
}

$requiredCommands = @(
    'Get-EntraAppSecrets',
    'Get-ExportDestination',
    'Export-Results'
)

$missingCommands = $requiredCommands | Where-Object {
    -not (Get-Command -Name $_ -ErrorAction SilentlyContinue)
}

if ($missingCommands) {
    throw (
        "Comandi richiesti non disponibili nella sessione dopo l'import dei moduli: {0}"
    ) -f ($missingCommands -join ', ')
}

if ([string]::IsNullOrWhiteSpace($SearchString)) {
    $SearchString = (Read-Host "Inserisci la stringa da cercare nel nome dell'app").Trim()

    if ([string]::IsNullOrWhiteSpace($SearchString)) {
        throw "Stringa di ricerca non specificata."
    }
}

if (-not $OutputPath) {
    $destination = Get-ExportDestination `
        -DefaultFileName "entra-app-secrets_$SearchString" `
        -Formats xlsx, csv `
        -PreferredFormat csv `
        -Title "Scegli dove salvare il report dei secret delle app Entra" `
        -UseConsole:$UseConsole `
        -Force:$Force

    $OutputPath = $destination.Path
}

Write-Verbose "Recupero dati delle app registrate..."
$results = Get-EntraAppSecrets `
    -SearchString $SearchString `
    -TenantId $TenantId `
    -UseDeviceCode:$UseDeviceCode `
    -AutoInstallModules:$AutoInstallModules `
    -ForceReconnect:$ForceReconnect `
    -PageSize $PageSize `
    -MaxResults $MaxResults `
    -Verbose:$VerbosePreference

$exportData = $results | Select-Object `
    @{ Name = 'Display Name';             Expression = { $_.DisplayName } },
    @{ Name = 'Application (Client) ID';  Expression = { $_.ApplicationClientId } },
    @{ Name = 'Object ID';                Expression = { $_.ObjectID } },
    @{ Name = 'Data Inizio Secret';       Expression = { $_.SecretStartDateTime } },
    @{ Name = 'Data Scadenza Secret';     Expression = { $_.SecretEndDateTime } },
    @{ Name = 'Secret ID';                Expression = { $_.SecretId } }

if ($PSCmdlet.ShouldProcess($OutputPath, "Esportazione report secret app Entra")) {
    Export-Results `
        -InputObject $exportData `
        -Path $OutputPath `
        -WorksheetName 'EntraAppSecrets' `
        -Force:$Force
}