[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string] $InputPath,

    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [ValidateSet('csv', 'xlsx')]
    [string] $ExportFormat = 'xlsx',

    [Parameter()]
    [string] $WorksheetName = 'TeamsMembersReport',

    [Parameter()]
    [string] $TeamIdColumn = 'TeamId',

    [Parameter()]
    [ValidateSet('UPN', 'MailNickName', 'DisplayName')]
    [string] $IdentityProperty = 'UPN',

    [Parameter()]
    [string] $Separator = ';',

    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [switch] $UseConsole,

    [Parameter()]
    [switch] $UseDeviceCode,

    [Parameter()]
    [switch] $AutoInstallModules,

    [Parameter()]
    [switch] $ForceReconnect,

    [Parameter()]
    [switch] $Force
)

begin {

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

    foreach ($module in $moduleCandidates.GetEnumerator()) {
        $moduleName = $module.Key
        $paths = $module.Value

        $found = $false
        foreach ($path in $paths) {
            if (Test-Path -Path $path -PathType Leaf) {
                Import-Module -Name $path -Force -ErrorAction Stop -Verbose:$false | Out-Null
                Write-Verbose "Modulo '$moduleName' importato da: $path"
                $found = $true
                break
            }
        }

        if (-not $found) {
            throw "Impossibile trovare il modulo '$moduleName'. Cercati nei seguenti percorsi:`n$($paths -join "`n")"
        }
    }

    $requiredCommands = @(
    'Get-TeamOwnersAndMembers',
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

    try {
        Write-Verbose "Inizializzazione script export Teams owners/members..."

        if (-not $InputPath) {
            Write-Verbose "Nessun file di input specificato. Avvio selezione file..."
            $InputPath = Get-InputFile `
                -Formats @('csv') `
                -Title 'Seleziona il file CSV contenente i Team ID' `
                -UseConsole:$UseConsole
        }

        if (-not (Test-Path $InputPath -PathType Leaf)) {
            throw "File di input non trovato: '$InputPath'"
        }

        $inputExtension = [IO.Path]::GetExtension($InputPath).ToLowerInvariant()
        if ($inputExtension -ne '.csv') {
            throw "Il file di input deve essere in formato CSV."
        }

        Write-Verbose "Import del file CSV di input: $InputPath"
        $inputRows = Import-Csv -Path $InputPath -ErrorAction Stop

        if (-not $inputRows) {
            throw "Il file di input non contiene righe."
        }

        if ($TeamIdColumn -notin $inputRows[0].PSObject.Properties.Name) {
            throw "La colonna '$TeamIdColumn' non è presente nel file CSV."
        }

        $teamIds = @(
            $inputRows |
            Select-Object -ExpandProperty $TeamIdColumn |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            ForEach-Object { $_.Trim() } |
            Sort-Object -Unique
        )

        if (-not $teamIds) {
            throw "Nessun Team ID valido trovato nella colonna '$TeamIdColumn'."
        }

        Write-Verbose "Totale Team ID univoci trovati: $($teamIds.Count)"

        if (-not $OutputPath) {
            Write-Verbose "Nessun file di output specificato. Avvio selezione destinazione export..."
            $defaultFileName = "TeamsOwnersMembersReport.$ExportFormat"
            $destination = Get-ExportDestination `
                -DefaultFileName $defaultFileName `
                -Formats @('xlsx', 'csv') `
                -PreferredFormat $ExportFormat `
                -Title 'Scegli dove salvare il report Teams Owners/Members' `
                -UseConsole:$UseConsole `
                -Force:$Force
            $OutputPath = $destination.Path
        }

        Write-Verbose "Percorso di output risolto: $OutputPath"
    }
    catch {
        throw
    }
}

process {
    try {
        $results = [System.Collections.Generic.List[object]]::new()
        $total = $teamIds.Count
        $index = 0

        Write-Verbose "Avvio recupero dati da Microsoft Graph per $total team..."

        foreach ($teamId in $teamIds) {
            $index++
            Write-Verbose ("[{0}/{1}] Elaborazione Team ID: {2}" -f $index, $total, $teamId)

            $currentResult = @(Get-TeamOwnersAndMembers `
                -TeamId $teamId `
                -IdentityProperty $IdentityProperty `
                -Separator $Separator `
                -TenantId $TenantId `
                -UseDeviceCode:$UseDeviceCode `
                -AutoInstallModules:$AutoInstallModules `
                -ForceReconnect:$ForceReconnect `
                -Verbose:$VerbosePreference)

            if ($currentResult) {
                foreach ($item in $currentResult) {
                    $results.Add($item)
                    Write-Verbose ("[{0}/{1}] Team elaborato: '{2}' | Owners: {3} | Members: {4}" -f `
                        $index,
                        $total,
                        $item.teamDisplayName,
                        $item.teamOwnersCount,
                        $item.teamMembersCount)
                }
            }
            else {
                Write-Verbose ("[{0}/{1}] Nessun dato restituito per Team ID: {2}" -f $index, $total, $teamId)
            }
        }

        if ($results.Count -eq 0) {
            throw "Nessun dato restituito dalla funzione Get-TeamOwnersAndMembers."
        }

        Write-Verbose "Recupero completato. Totale record prodotti: $($results.Count)"
        Write-Verbose "Avvio esportazione del report..."

        if ($PSCmdlet.ShouldProcess($OutputPath, 'Esportazione report Teams Owners/Members')) {
            $exportResult = Export-Results `
                -InputObject $results.ToArray() `
                -Path $OutputPath `
                -WorksheetName $WorksheetName `
                -Force:$Force

            Write-Verbose "Esportazione completata. File generato: $($exportResult.Path)"
            Write-Host ""
            Write-Host "Report esportato correttamente:" -ForegroundColor Green
            Write-Host $exportResult.Path
            Write-Host "Formato:" $exportResult.Format
        }
    }
    catch {
        throw
    }
}