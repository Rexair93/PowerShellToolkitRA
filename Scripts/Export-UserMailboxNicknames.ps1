<#
.SYNOPSIS
    Recupera i MailNickName degli utenti da un CSV ed esporta i risultati.

.DESCRIPTION
    Legge un file CSV con colonna 'UPN', interroga Microsoft Entra tramite
    Get-UserMailboxNickname e salva l'output in formato XLSX o CSV.
    Se InputPath o OutputPath non vengono specificati, vengono richiesti
    tramite Get-InputFile / Get-ExportDestination (GUI o console).

.PARAMETER InputPath
    Percorso del file CSV di input. Se omesso, viene richiesto interattivamente.

.PARAMETER OutputPath
    Percorso completo del file di output. Se omesso, viene richiesto interattivamente.

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
    .\Export-UserMailboxNicknames.ps1 -InputPath .\users.csv

.EXAMPLE
    .\Export-UserMailboxNicknames.ps1 -InputPath .\users.csv -OutputPath .\out.xlsx -UseConsole
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string] $InputPath,

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
    [switch] $AllowClobber,

    [Parameter()]
    [switch] $ForceReconnect,

    [Parameter()]
    [switch] $Force
)

# ---------------------------------------------------------------------------
# Import moduli necessari
# ---------------------------------------------------------------------------
Import-Module CloudOperations  -ErrorAction Stop
Import-Module FilesUtilities   -ErrorAction Stop

# ---------------------------------------------------------------------------
try {
    # -----------------------------------------------------------------------
    # 1. Risolvi il file di input
    # -----------------------------------------------------------------------
    if (-not $InputPath) {
        Write-Verbose "Richiesta file di input..."
        $InputPath = Get-InputFile `
            -Formats    @('csv') `
            -Title      'Seleziona il CSV con gli UPN degli utenti' `
            -UseConsole:$UseConsole
    }

    # -----------------------------------------------------------------------
    # 2. Valida il CSV immediatamente (fallisce velocemente)
    # -----------------------------------------------------------------------
    if (-not (Test-Path $InputPath -PathType Leaf)) {
        throw "Il file di input '$InputPath' non esiste."
    }

    Write-Verbose "Lettura e validazione file CSV..."
    $inputRows = Import-Csv -Path $InputPath -ErrorAction Stop

    if (-not $inputRows) {
        throw "Il file CSV di input è vuoto."
    }

    if ('UPN' -notin $inputRows[0].PSObject.Properties.Name) {
        throw "Il file CSV non contiene la colonna 'UPN'."
    }

    # -----------------------------------------------------------------------
    # 3. Risolvi il percorso di output
    # -----------------------------------------------------------------------
    if (-not $OutputPath) {
        Write-Verbose "Richiesta percorso di output..."
        $destination = Get-ExportDestination `
            -DefaultFileName  'UsersMailNicknames.xlsx' `
            -Formats          @('xlsx', 'csv') `
            -PreferredFormat  'xlsx' `
            -Title            "Scegli dove salvare l'export dei MailNickName utenti" `
            -UseConsole:$UseConsole `
            -Force:$Force
        $OutputPath = $destination.Path
    }

    # -----------------------------------------------------------------------
    # 4. Recupera i dati da Entra
    # -----------------------------------------------------------------------
    Write-Verbose "Avvio recupero MailNickName utenti..."
    $results = $inputRows |
        Select-Object -ExpandProperty UPN |
        Get-UserMailboxNicknames `
            -TenantId           $TenantId `
            -UseDeviceCode:     $UseDeviceCode `
            -ForceReconnect:    $ForceReconnect `
            -AutoInstallModules:$AutoInstallModules `
            -AllowClobber:      $AllowClobber `
            -Verbose:           $VerbosePreference |
        Sort-Object UserPrincipalName

    if (-not $results) {
        Write-Warning "Nessun risultato da esportare."
        exit 0
    }

    # -----------------------------------------------------------------------
    # 5. Esporta i risultati
    # -----------------------------------------------------------------------
    if ($PSCmdlet.ShouldProcess($OutputPath, "Esporta MailNickName utenti")) {
        Write-Verbose "Esportazione in corso verso '$OutputPath'..."
        $export = Export-Results `
            -InputObject    $results `
            -Path           $OutputPath `
            -WorksheetName  'UsersMailNicknames' `
            -Force:$Force

        Write-Host "Esportazione completata: $($export.Path) [$($export.Format.ToUpper())]" -ForegroundColor Green
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