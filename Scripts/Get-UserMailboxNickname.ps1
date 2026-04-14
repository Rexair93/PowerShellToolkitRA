#Requires -Version 7.0
<#
.SYNOPSIS
    Recupera MailNickName degli utenti Entra a partire da un CSV con colonna UPN.

.DESCRIPTION
    Usa i moduli personalizzati:
    - CloudConnect     -> Connect-ToEntra
    - FilesUtilities   -> Get-InputFile, Get-ExportDestination, Export-Results

.PARAMETER TenantId
    Tenant ID opzionale.

.PARAMETER UseDeviceCode
    Usa autenticazione device code, se supportata dal modulo.

.PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti.

.PARAMETER UseConsole
    Forza input/output in modalità console.

.PARAMETER Force
    Consente la sovrascrittura del file di output.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [switch] $UseDeviceCode,

    [Parameter()]
    [switch] $AutoInstallModules,

    [Parameter()]
    [switch] $UseConsole,

    [Parameter()]
    [switch] $Force
)

$ErrorActionPreference = 'Stop'

#region Import moduli personalizzati
Import-Module CloudConnect   -ErrorAction Stop
Import-Module FilesUtilities -ErrorAction Stop
#endregion

#region Connessione a Entra
$connectParams = @{
    AutoInstallModules = $AutoInstallModules
    UseDeviceCode      = $UseDeviceCode
}
if ($TenantId) {
    $connectParams.TenantId = $TenantId
}

Connect-ToEntra @connectParams
#endregion

#region Selezione file input
$inputPath = Get-InputFile -Formats @('csv') -Title 'Seleziona il file CSV di input con colonna UPN' -UseConsole:$UseConsole
#endregion

#region Import e validazione CSV
$rows = Import-Csv -Path $inputPath

if (-not $rows) {
    throw "Il file CSV '$inputPath' è vuoto o non contiene righe valide."
}

$firstRow = $rows | Select-Object -First 1
if (-not ($firstRow.PSObject.Properties.Name -contains 'UPN')) {
    throw "Il file CSV deve contenere una colonna chiamata 'UPN'."
}

$upnList = $rows |
    Select-Object -ExpandProperty UPN |
    Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
#endregion

#region Elaborazione utenti
$outputData = foreach ($userUPN in $upnList) {
    $escapedUPN = $userUPN -replace "'", "''"

    try {
        $user = Get-EntraUser -Filter "UserPrincipalName eq '$escapedUPN'" -ErrorAction Stop

        if (-not $user) {
            [pscustomobject]@{
                ObjectId     = $null
                UPN          = $userUPN
                MailNickName = $null
                Status       = 'UserNotFound'
            }
            continue
        }

        [pscustomobject]@{
            ObjectId     = $user.Id
            UPN          = $user.UserPrincipalName
            MailNickName = $user.MailNickname
            Status       = 'OK'
        }
    }
    catch {
        [pscustomobject]@{
            ObjectId     = $null
            UPN          = $userUPN
            MailNickName = $null
            Status       = $_.Exception.Message
        }
    }
}
#endregion

#region Selezione destinazione output
$destination = Get-ExportDestination -DefaultFileName 'UsersMailNicknames.xlsx' -Formats @('xlsx', 'csv') -PreferredFormat 'xlsx' -Title 'Scegli dove salvare il report' -UseConsole:$UseConsole -Force:$Force
#endregion

#region Export risultati
$result = Export-Results -InputObject $outputData -Path $destination.Path -WorksheetName 'UsersMailNicknames' -Force:$Force
#endregion

Write-Host "Esportazione completata: $($result.Path) [$($result.Format)]" -ForegroundColor Green