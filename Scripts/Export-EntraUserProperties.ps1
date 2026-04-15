<#
.SYNOPSIS
    Recupera proprietà arbitrarie di utenti Entra ID da un CSV ed esporta i risultati.

.DESCRIPTION
    Legge un file CSV contenente una colonna configurabile (es. 'EmployeeId', 'UPN'...),
    interroga Microsoft Entra tramite Get-EntraUserProperties e salva l'output in
    formato XLSX o CSV.
    Se InputPath o OutputPath non vengono specificati, vengono richiesti
    tramite Get-InputFile / Get-ExportDestination (GUI o console).

.PARAMETER InputPath
    Percorso del file CSV di input. Se omesso, viene richiesto interattivamente.

.PARAMETER InputColumn
    Nome della colonna del CSV che contiene i valori di ricerca.
    Default: 'UserPrincipalName'.

.PARAMETER LookupProperty
    Proprietà Entra ID su cui costruire il filtro OData.
    Default: 'UserPrincipalName'.

.PARAMETER Properties
    Elenco delle proprietà Entra da includere nell'output.
    Default: 'ObjectId', 'UserPrincipalName', 'MailNickName'.

.PARAMETER OutputPath
    Percorso completo del file di output. Se omesso, viene richiesto interattivamente.

.PARAMETER WorksheetName
    Nome del foglio Excel. Default: 'EntraUserProperties'.

.PARAMETER TenantId
    Tenant ID da usare per la connessione a Microsoft Entra.

.PARAMETER UseDeviceCode
    Usa l'autenticazione device code se supportata.

.PARAMETER UseConsole
    Forza l'uso della modalità console per la selezione dei percorsi.

.PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti.

.PARAMETER AllowClobber
    Consente AllowClobber durante l'installazione del modulo Microsoft.Entra.

.PARAMETER ForceReconnect
    Forza una nuova connessione a Microsoft Entra.

.PARAMETER Force
    Sovrascrive il file di output se già esistente.

.EXAMPLE
    .\Export-EntraUserProperties.ps1 -InputPath .\users.csv

.EXAMPLE
    # Ricerca per EmployeeId, recupera DisplayName e Mail
    .\Export-EntraUserProperties.ps1 `
        -InputPath      .\employees.csv `
        -InputColumn    'EmployeeId' `
        -LookupProperty 'EmployeeId' `
        -Properties     'UserPrincipalName','ObjectId','DisplayName','Mail' `
        -OutputPath     .\output\EntraUsers.xlsx

.EXAMPLE
    .\Export-EntraUserProperties.ps1 -InputPath .\users.csv -OutputPath .\out.xlsx -UseConsole -Verbose
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string] $LookupProperty = 'UserPrincipalName',

    [Parameter()]
    [string[]] $Properties = @('ObjectId', 'UserPrincipalName', 'GivenName', 'Surname', 'DisplayName', 'MailNickName', 'EmployeeId'),

    [Parameter()]
    [string] $InputPath,

    [Parameter()]
    [string] $InputColumn = $LookupProperty,

    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [string] $WorksheetName = 'EntraUserProperties',

    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [switch] $UseDeviceCode,

    [Parameter()]
    [switch] $UseConsole,

    [Parameter()]
    [switch] $AutoInstallModules,

    [Parameter()]
    [switch] $AllowClobber,

    [Parameter()]
    [switch] $ForceReconnect,

    [Parameter()]
    [switch] $Force
)

# ---------------------------------------------------------------------------
# Import moduli necessari
# ---------------------------------------------------------------------------
Import-Module CloudOperations -ErrorAction Stop
Import-Module FilesUtilities  -ErrorAction Stop

# ---------------------------------------------------------------------------
try {
    # -----------------------------------------------------------------------
    # 1. Risolvi il file di input
    # -----------------------------------------------------------------------
    if (-not $InputPath) {
        Write-Verbose "Richiesta file di input..."
        $InputPath = Get-InputFile `
            -Formats    @('csv') `
            -Title      "Seleziona il CSV con i valori '$InputColumn'" `
            -UseConsole:$UseConsole
    }

    # -----------------------------------------------------------------------
    # 2. Valida il CSV (fail fast)
    # -----------------------------------------------------------------------
    if (-not (Test-Path $InputPath -PathType Leaf)) {
        throw "Il file di input '$InputPath' non esiste."
    }

    Write-Verbose "Lettura e validazione file CSV..."
    $inputRows = Import-Csv -Path $InputPath -ErrorAction Stop

    if (-not $inputRows) {
        throw "Il file CSV di input è vuoto."
    }

    $availableColumns = $inputRows[0].PSObject.Properties.Name
    if ($InputColumn -notin $availableColumns) {
        $colList = $availableColumns -join ', '
        throw "Il file CSV non contiene la colonna '$InputColumn'. Colonne disponibili: $colList"
    }

    # -----------------------------------------------------------------------
    # 3. Risolvi il percorso di output
    # -----------------------------------------------------------------------
    if (-not $OutputPath) {
        Write-Verbose "Richiesta percorso di output..."
        $destination = Get-ExportDestination `
            -DefaultFileName "EntraUserProperties.xlsx" `
            -Formats         @('xlsx', 'csv') `
            -PreferredFormat 'xlsx' `
            -Title           "Scegli dove salvare l'export delle proprietà utente" `
            -UseConsole:     $UseConsole `
            -Force:          $Force
        $OutputPath = $destination.Path
    }

    # -----------------------------------------------------------------------
    # 4. Recupera i dati da Entra
    # -----------------------------------------------------------------------
    Write-Verbose "Avvio recupero proprietà utenti (lookup: $LookupProperty)..."
    $results = $inputRows |
    Select-Object -ExpandProperty $InputColumn |
    Get-EntraUserProperties `
        -LookupProperty      $LookupProperty `
        -Properties          $Properties `
        -TenantId            $TenantId `
        -UseDeviceCode:      $UseDeviceCode `
        -ForceReconnect:     $ForceReconnect `
        -AutoInstallModules: $AutoInstallModules `
        -AllowClobber:       $AllowClobber `
        -Verbose:            $VerbosePreference |
    Sort-Object LookupValue

    if (-not $results) {
        Write-Warning "Nessun risultato da esportare."
        exit 0
    }

    $foundCount = ($results | Where-Object Found).Count
    $notFoundCount = ($results | Where-Object { -not $_.Found }).Count
    Write-Verbose "Utenti trovati: $foundCount | Non trovati: $notFoundCount"

    # -----------------------------------------------------------------------
    # 5. Esporta i risultati
    # -----------------------------------------------------------------------
    if ($PSCmdlet.ShouldProcess($OutputPath, "Esporta proprietà utenti Entra")) {
        Write-Verbose "Esportazione in corso verso '$OutputPath'..."
        $export = Export-Results `
            -InputObject   $results `
            -Path          $OutputPath `
            -WorksheetName $WorksheetName `
            -Force:        $Force

        Write-Host "Esportazione completata: $($export.Path) [$($export.Format.ToUpper())]" `
            -ForegroundColor Green
        Write-Host "  Trovati:     $foundCount" -ForegroundColor Cyan
        if ($notFoundCount -gt 0) {
            Write-Host "  Non trovati: $notFoundCount" -ForegroundColor Yellow
        }
    }
}
catch {
    # -----------------------------------------------------------------------
    # Gestione centralizzata degli errori
    # -----------------------------------------------------------------------
    if ($_.Exception.Message -eq 'Operazione annullata.') {
        Write-Warning "Operazione annullata dall'utente."
        exit 0
    }

    Write-Error $_.Exception.Message
    exit 1
}
