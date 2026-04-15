<#
.SYNOPSIS
    Esporta i record DNS SRV _sipfederationtls._tcp in XLSX o CSV.

.DESCRIPTION
    Con -AllDomains recupera tutti i domini del tenant tramite Microsoft Graph.
    Senza -AllDomains (default) legge l'elenco domini da un file CSV o Excel
    con colonna 'Domain', specificato con -InputPath o selezionato interattivamente.

.PARAMETER AllDomains
    Se specificato, interroga tutti i domini del tenant. Mutualmente esclusivo con -InputPath.

.PARAMETER InputPath
    Percorso di un file CSV o Excel con colonna 'Domain'. Se omesso, viene richiesto
    interattivamente. Mutualmente esclusivo con -AllDomains.

.PARAMETER OutputPath
    Percorso del file di output. Se omesso, viene richiesto interattivamente.

.PARAMETER Scopes
    Scope Microsoft Graph. Default: 'Domain.Read.All'.

.PARAMETER TenantId
    Tenant ID per la connessione a Microsoft Graph.

.PARAMETER UseDeviceCode
    Usa l'autenticazione device code.

.PARAMETER UseConsole
    Forza la selezione dei percorsi in modalità console.

.PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti.

.PARAMETER ForceReconnect
    Forza una nuova connessione a Microsoft Graph.

.PARAMETER DnsServer
    Server DNS specifico per la risoluzione SRV.

.PARAMETER Force
    Sovrascrive il file di output se già esistente.

.EXAMPLE
    # Tutti i domini del tenant
    .\Export-SipFederationTlsSrvRecords.ps1 -AllDomains

.EXAMPLE
    # Solo i domini in un file CSV
    .\Export-SipFederationTlsSrvRecords.ps1 -InputPath .\domains.csv

.EXAMPLE
    # Selezione interattiva del file di input
    .\Export-SipFederationTlsSrvRecords.ps1
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Filtered')]
param(
    [Parameter(ParameterSetName = 'All', Mandatory)]
    [switch] $AllDomains,

    [Parameter(ParameterSetName = 'Filtered')]
    [string] $InputPath,

    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [string[]] $Scopes = @('Domain.ReadWrite.All'),

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
    [string] $DnsServer,

    [Parameter()]
    [switch] $Force
)

Import-Module CloudOperations -ErrorAction Stop
Import-Module DnsTools        -ErrorAction Stop
Import-Module FilesUtilities  -ErrorAction Stop

try {
    # -----------------------------------------------------------------------
    # 1. Determina i domini da interrogare
    # -----------------------------------------------------------------------
    $domainsToQuery = $null   # $null → Get-SipFederationTlsSrvRecords userà Get-MgDomain

    if (-not $AllDomains) {
        # Risolvi il file di input
        if (-not $InputPath) {
            Write-Verbose "Richiesta file di input..."
            $InputPath = Get-InputFile `
                -Formats    @('csv', 'xlsx') `
                -Title      'Seleziona il file con i domini (colonna "Domain")' `
                -UseConsole:$UseConsole
        }

        if (-not (Test-Path $InputPath -PathType Leaf)) {
            throw "Il file di input '$InputPath' non esiste."
        }

        # Leggi CSV o Excel in base all'estensione
        Write-Verbose "Lettura file di input '$InputPath'..."
        $ext = [IO.Path]::GetExtension($InputPath).TrimStart('.').ToLowerInvariant()

        $inputRows = switch ($ext) {
            'csv'  { Import-Csv   -Path $InputPath -ErrorAction Stop }
            'xlsx' { Import-Excel -Path $InputPath -ErrorAction Stop }
            default { throw "Formato '$ext' non supportato. Usa CSV o XLSX." }
        }

        if (-not $inputRows) {
            throw "Il file di input è vuoto."
        }

        if ('Domain' -notin $inputRows[0].PSObject.Properties.Name) {
            throw "Il file di input non contiene la colonna 'Domain'."
        }

        $domainsToQuery = $inputRows |
            Select-Object -ExpandProperty Domain |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

        if (-not $domainsToQuery) {
            throw "Nessun dominio valido trovato nel file di input."
        }

        Write-Verbose "$($domainsToQuery.Count) domini letti dal file."
    }

    # -----------------------------------------------------------------------
    # 2. Recupera i record SRV
    # -----------------------------------------------------------------------
    $getParams = @{
        Scopes             = $Scopes
        TenantId           = $TenantId
        UseDeviceCode      = $UseDeviceCode.IsPresent
        AutoInstallModules = $AutoInstallModules.IsPresent
        ForceReconnect     = $ForceReconnect.IsPresent
        Verbose            = $VerbosePreference
    }
    if ($DnsServer) { $getParams.DnsServer = $DnsServer }

    $results = if ($domainsToQuery) {
        $domainsToQuery | Get-SipFederationTlsSrvRecords @getParams
    } else {
        Get-SipFederationTlsSrvRecords @getParams
    }

    if (-not $results) {
        exit 0
    }

    # -----------------------------------------------------------------------
    # 3. Risolvi il percorso di output
    # -----------------------------------------------------------------------
    if (-not $OutputPath) {
        Write-Verbose "Richiesta percorso di output..."
        $destination = Get-ExportDestination `
            -DefaultFileName "sipfederationtls_srv_records.xlsx" `
            -Formats         @('xlsx', 'csv') `
            -PreferredFormat 'xlsx' `
            -Title           "Scegli dove salvare i record SRV sipfederationtls" `
            -UseConsole:     $UseConsole `
            -Force:          $Force
        $OutputPath = $destination.Path
    }

    # -----------------------------------------------------------------------
    # 4. Esporta
    # -----------------------------------------------------------------------
    if ($PSCmdlet.ShouldProcess($OutputPath, "Esporta record SRV sipfederationtls")) {
        Write-Verbose "Esportazione in corso verso '$OutputPath'..."
        $export = Export-Results `
            -InputObject   $results `
            -Path          $OutputPath `
            -WorksheetName 'SRV_sipfederationtls' `
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