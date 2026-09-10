<#
.SYNOPSIS
Aggiunge utenti a Team e canali privati Microsoft Teams da un file CSV o XLSX.

.DESCRIPTION
Utilizza Add-TeamsMembershipFromSpreadsheet del modulo CloudOperations per leggere
una lista di membership, risolvere le identità utente e aggiungere gli utenti alla
destinazione indicata nel file.

Accetta anche direttamente il report prodotto da
Export-UserTeamsAndPrivateChannelsReport.ps1, mantenendo i nomi delle colonne
UserPrincipalName, TeamRole, ChannelRole e ChannelId.

Se -Path non viene specificato, il file viene selezionato tramite Get-InputFile.
Al termine viene esportato un report XLSX tramite Export-Results.

.PARAMETER Path
Percorso del file CSV o XLSX contenente la lista delle membership.

.PARAMETER WorksheetName
Nome del foglio da leggere, se il file è XLSX.

.PARAMETER Delimiter
Delimitatore CSV. Default: ','.

.PARAMETER OutputPath
Percorso del report XLSX. Se omesso, viene richiesto all'utente.

.PARAMETER OutputWorksheetName
Nome del foglio del report. Default: Results.

.PARAMETER UserColumn
Nome della colonna contenente l'identità utente. Default: User.

.PARAMETER TargetTypeColumn
Nome della colonna contenente il tipo di destinazione. Default: TargetType.

.PARAMETER TeamIdColumn
Nome della colonna contenente il Team ID. Default: TeamId.

.PARAMETER TeamDisplayNameColumn
Nome della colonna contenente il nome del Team. Default: TeamDisplayName.

.PARAMETER ChannelDisplayNameColumn
Nome della colonna contenente il nome del canale privato. Default: ChannelDisplayName.

.PARAMETER RoleColumn
Nome della colonna contenente il ruolo. Default: Role.

.PARAMETER IdentityType
Tipo di identità in input: Auto, UserPrincipalName, ObjectId, MailNickName.

.PARAMETER TenantId
Tenant ID da usare per la connessione ai servizi Microsoft.

.PARAMETER UseDeviceCode
Usa autenticazione device code.

.PARAMETER AutoInstallModules
Installa automaticamente i moduli mancanti richiesti.

.PARAMETER AllowClobber
Consente AllowClobber durante l'installazione dei moduli.

.PARAMETER ForceReconnect
Forza una nuova connessione ai servizi richiesti.

.PARAMETER UseConsole
Usa la console per la selezione dei file invece dell'interfaccia grafica.

.EXAMPLE
.\Add-TeamsMembershipFromSpreadsheet.ps1 -Path 'C:\Temp\members.xlsx' -Verbose

.EXAMPLE
.\Add-TeamsMembershipFromSpreadsheet.ps1 -UseConsole -Delimiter ';'
#>
[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    [Parameter()]
    [string] $Path,

    [Parameter()]
    [string] $WorksheetName,

    [Parameter()]
    [char] $Delimiter = ',',

    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [string] $OutputWorksheetName = 'Results',

    [Parameter()]
    [string] $UserColumn = 'User',

    [Parameter()]
    [string] $TargetTypeColumn = 'TargetType',

    [Parameter()]
    [string] $TeamIdColumn = 'TeamId',

    [Parameter()]
    [string] $TeamDisplayNameColumn = 'TeamDisplayName',

    [Parameter()]
    [string] $ChannelDisplayNameColumn = 'ChannelDisplayName',

    [Parameter()]
    [string] $RoleColumn = 'Role',

    [Parameter()]
    [ValidateSet('Auto', 'UserPrincipalName', 'ObjectId', 'MailNickName')]
    [string] $IdentityType = 'Auto',

    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [switch] $UseDeviceCode,

    [Parameter()]
    [switch] $AutoInstallModules,

    [Parameter()]
    [switch] $AllowClobber,

    [Parameter()]
    [switch] $ForceReconnect,

    [Parameter()]
    [switch] $UseConsole
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot

$cloudOperationsManifest = Join-Path $repoRoot 'Modules\CloudOperations\CloudOperations.psd1'
$filesUtilitiesManifest  = Join-Path $repoRoot 'Modules\FilesUtilities\FilesUtilities.psd1'

try {
    Import-Module $filesUtilitiesManifest -Force -ErrorAction Stop
    Import-Module $cloudOperationsManifest -Force -ErrorAction Stop

    if ([string]::IsNullOrWhiteSpace($Path)) {
        $Path = Get-InputFile `
            -Formats @('csv', 'xlsx') `
            -Title 'Seleziona il file con la lista delle membership Teams' `
            -UseConsole:$UseConsole
    }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw 'Nessun file di input selezionato.'
    }

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $defaultFileName = "Add-TeamsMembershipFromSpreadsheet-$timestamp.xlsx"

        $OutputPath = Get-ExportDestination `
            -DefaultFileName $defaultFileName `
            -InitialDirectory (Split-Path -Parent $Path) `
            -Formats @('xlsx') `
            -PreferredFormat 'xlsx' `
            -Title 'Scegli dove salvare il report Excel' `
            -UseConsole:$UseConsole `
            -Force `
            -AsString
    }

    $results = @(
        Add-TeamsMembershipFromSpreadsheet `
            -Path $Path `
            -WorksheetName $WorksheetName `
            -Delimiter $Delimiter `
            -UserColumn $UserColumn `
            -TargetTypeColumn $TargetTypeColumn `
            -TeamIdColumn $TeamIdColumn `
            -TeamDisplayNameColumn $TeamDisplayNameColumn `
            -ChannelDisplayNameColumn $ChannelDisplayNameColumn `
            -RoleColumn $RoleColumn `
            -IdentityType $IdentityType `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -AutoInstallModules:$AutoInstallModules `
            -AllowClobber:$AllowClobber `
            -ForceReconnect:$ForceReconnect `
            -WhatIf:$WhatIfPreference `
            -Verbose:$VerbosePreference
    )

    $exportInfo = $results |
        Export-Results `
            -Path $OutputPath `
            -WorksheetName $OutputWorksheetName `
            -Force

    $successCount = @($results | Where-Object Success).Count
    $errorCount = @($results | Where-Object { -not $_.Success }).Count

    Write-Host ("Operazioni completate. Riuscite: {0}. Errori: {1}. Report: {2}" -f $successCount, $errorCount, $exportInfo.Path)
}
catch {
    Write-Error $_
    exit 1
}
