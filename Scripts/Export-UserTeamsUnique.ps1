    <#
    .SYNOPSIS
        Esporta l'elenco univoco dei team Microsoft Teams associati a una lista di utenti.

    .DESCRIPTION
        Legge un file CSV o XLSX di input contenente una colonna con UPN o MailNickName,
        recupera i team di appartenenza per ciascun utente tramite Get-UserTeamMemberships
        e produce un file di export contenente ogni team una sola volta.

        La deduplicazione avviene in tempo reale durante il recupero: ogni team già visto
        viene scartato prima ancora di essere emesso in pipeline da Get-UserTeamMemberships.

        Per default, sia il file di input sia il file di output vengono selezionati
        tramite finestra di dialogo. Se -UseConsole è specificato oppure la GUI non
        è disponibile, viene usata la console.

    .PARAMETER InputPath
        Percorso del file CSV o XLSX di input. Se non specificato, viene richiesto
        tramite Get-InputFile.

    .PARAMETER InputColumn
        Nome della colonna contenente i valori utente.
        Default: Identity

    .PARAMETER InputWorksheetName
        Nome del foglio Excel da leggere quando il file di input è .xlsx.
        Se omesso, viene usato il primo foglio disponibile.

    .PARAMETER IdentityType
        Tipo di identificativo usato nella colonna InputColumn.
        Valori ammessi: Auto, UserPrincipalName, MailNickName.
        Default: Auto

    .PARAMETER OutputPath
        Percorso completo del file di output. Se non specificato, viene richiesto
        tramite Get-ExportDestination.

    .PARAMETER TenantId
        Tenant ID da usare per la connessione ai servizi Microsoft.

    .PARAMETER UseDeviceCode
        Prova a usare l'autenticazione device code.

    .PARAMETER UseConsole
        Forza la selezione dei percorsi in modalità console.

    .PARAMETER AutoInstallModules
        Installa automaticamente i moduli mancanti richiesti.

    .PARAMETER AllowClobber
        Consente l'uso di AllowClobber durante l'installazione dei moduli richiesti.

    .PARAMETER ForceReconnect
        Forza una nuova connessione ai servizi richiesti.

    .PARAMETER Force
        Consente la sovrascrittura del file di output quando supportato.

    .OUTPUTS
        PSCustomObject con proprietà Path e Format.

    .EXAMPLE
        Export-UserTeamsUnique

    .EXAMPLE
        Export-UserTeamsUnique -InputPath .\utenti.csv -InputColumn UPN -OutputPath .\teams.xlsx -Verbose
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string] $InputPath,

        [Parameter()]
        [string] $InputColumn = 'Identity',

        [Parameter()]
        [string] $InputWorksheetName,

        [Parameter()]
        [ValidateSet('Auto', 'UserPrincipalName', 'MailNickName')]
        [string] $IdentityType = 'Auto',

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
    'Get-UserTeamMemberships',
    'Get-InputFile',
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

    if (-not $InputPath) {
        $InputPath = Get-InputFile `
            -Formats    csv, xlsx `
            -Title      "Seleziona il file di input con gli utenti" `
            -UseConsole:$UseConsole
    }

    if (-not $OutputPath) {
        $OutputPath = Get-ExportDestination `
            -DefaultFileName "teams-univoci.xlsx" `
            -Formats         xlsx, csv `
            -PreferredFormat xlsx `
            -Title           "Scegli dove salvare l'export dei team univoci" `
            -UseConsole:     $UseConsole `
            -Force:          $Force `
            -AsString
    }

    $inputExt = ([IO.Path]::GetExtension($InputPath)).ToLowerInvariant()

    Write-Verbose "Import del file di input '$InputPath'..."
    switch ($inputExt) {
        '.csv' {
            $rows = Import-CsvSafe -Path $InputPath
        }

        '.xlsx' {
            if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                throw "Per leggere file Excel in input è richiesto il modulo 'ImportExcel'."
            }

            if ([string]::IsNullOrWhiteSpace($InputWorksheetName)) {
                $sheetInfo = Get-ExcelSheetInfo -Path $InputPath -ErrorAction Stop |
                    Select-Object -First 1

                if (-not $sheetInfo) {
                    throw "Il file Excel di input non contiene fogli leggibili."
                }

                $InputWorksheetName = $sheetInfo.Name
            }

            $rows = Import-Excel -Path $InputPath -WorksheetName $InputWorksheetName -ErrorAction Stop
        }

        default {
            throw "Formato di input non supportato: '$inputExt'."
        }
    }

    if (-not $rows) {
        throw "Il file di input non contiene righe utili."
    }

    if ($InputColumn -notin $rows[0].PSObject.Properties.Name) {
        $available = ($rows[0].PSObject.Properties.Name) -join ', '
        throw "La colonna '$InputColumn' non è presente nel file di input. Colonne disponibili: $available"
    }

    $identities = $rows |
        Select-Object -ExpandProperty $InputColumn |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        ForEach-Object { "$_".Trim() } |
        Sort-Object -Unique

    if (-not $identities) {
        throw "Nessun valore valido trovato nella colonna '$InputColumn'."
    }

    Write-Verbose "Recupero team univoci per $($identities.Count) utenti..."
    $uniqueTeams = @(
        $identities | Get-UserTeamMemberships `
            -IdentityType        $IdentityType `
            -TenantId            $TenantId `
            -UseDeviceCode:      $UseDeviceCode `
            -AutoInstallModules: $AutoInstallModules `
            -AllowClobber:       $AllowClobber `
            -ForceReconnect:     $ForceReconnect `
            -Verbose:            $VerbosePreference
    )

    if (-not $uniqueTeams) {
        Write-Warning "Nessun team trovato per gli utenti specificati."
        $uniqueTeams = @()
    }

    Write-Verbose "Team univoci trovati: $($uniqueTeams.Count). Avvio export..."
    Export-Results `
        -InputObject   $uniqueTeams `
        -Path          $OutputPath `
        -WorksheetName "Teams" `
        -Force:        $Force