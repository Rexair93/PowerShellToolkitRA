[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string] $TeamId,

    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [string] $WorksheetName,

    [Parameter()]
    [string] $IdentityColumn,

    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [string] $OutputWorksheetName = 'Results',

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

Import-Module $filesUtilitiesManifest -Force -ErrorAction Stop
Import-Module $cloudOperationsManifest -Force -ErrorAction Stop

$inputPath = Get-InputFile `
    -Formats @('csv', 'xlsx') `
    -Title 'Seleziona il file con la lista utenti da aggiungere al Team' `
    -UseConsole:$UseConsole

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $defaultFileName = "Add-TeamMembersFromIdentityList-$timestamp.xlsx"

    $OutputPath = Get-ExportDestination `
        -DefaultFileName $defaultFileName `
        -InitialDirectory (Split-Path -Parent $inputPath) `
        -Formats @('xlsx') `
        -PreferredFormat 'xlsx' `
        -Title 'Scegli dove salvare il report Excel' `
        -UseConsole:$UseConsole `
        -Force `
        -AsString
}

$identityRows = @(Import-TeamsUserIdentityList `
    -Path $inputPath `
    -WorksheetName $WorksheetName `
    -IdentityColumn $IdentityColumn)

if (-not $identityRows) {
    throw "Nessuna identità trovata nel file specificato."
}

$results = [System.Collections.Generic.List[object]]::new()

foreach ($identityRow in $identityRows) {
    $resolvedUser = $null

    try {
        $resolvedUser = @(
            Resolve-EntraUserIdentity `
                -Identity $identityRow.SourceIdentity `
                -IdentityType $identityRow.SourceIdentityType `
                -TenantId $TenantId `
                -UseDeviceCode:$UseDeviceCode `
                -AutoInstallModules:$AutoInstallModules `
                -AllowClobber:$AllowClobber `
                -ForceReconnect:$ForceReconnect `
                -Verbose:$VerbosePreference
        ) | Select-Object -First 1
    }
    catch {
        $results.Add(
            [pscustomobject][ordered]@{
                RowNumber          = $identityRow.RowNumber
                SourceIdentity     = $identityRow.SourceIdentity
                SourceIdentityType = $identityRow.SourceIdentityType
                TeamId             = $TeamId
                TeamDisplayName    = $null
                UserObjectId       = $null
                UserPrincipalName  = $null
                UserDisplayName    = $null
                UserMailNickName   = $null
                Status             = 'ResolveFailed'
                Message            = $_.Exception.Message
            }
        )
        continue
    }

    if ($null -eq $resolvedUser) {
        $results.Add(
            [pscustomobject][ordered]@{
                RowNumber          = $identityRow.RowNumber
                SourceIdentity     = $identityRow.SourceIdentity
                SourceIdentityType = $identityRow.SourceIdentityType
                TeamId             = $TeamId
                TeamDisplayName    = $null
                UserObjectId       = $null
                UserPrincipalName  = $null
                UserDisplayName    = $null
                UserMailNickName   = $null
                Status             = 'ResolveFailed'
                Message            = 'Errore durante la risoluzione dell''identità.'
            }
        )
        continue
    }

    $requestedRole = if ($identityRow.PSObject.Properties['RequestedRole']) {
        $identityRow.RequestedRole
    }
    else {
        'Member'
    }

    $addResults = @(
        $resolvedUser | Add-TeamGroupMemberSafe `
            -TeamId $TeamId `
            -Role $requestedRole `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -AutoInstallModules:$AutoInstallModules `
            -ForceReconnect:$ForceReconnect `
            -Verbose:$VerbosePreference
    )

    foreach ($result in $addResults) {
        $results.Add(
            [pscustomobject][ordered]@{
                RowNumber          = $identityRow.RowNumber
                SourceIdentity     = $identityRow.SourceIdentity
                SourceIdentityType = $identityRow.SourceIdentityType
                RequestedRole      = $requestedRole
                EffectiveRole      = $result.EffectiveRole
                TeamId             = $result.TeamId
                TeamDisplayName    = $result.TeamDisplayName
                UserObjectId       = $result.UserObjectId
                UserPrincipalName  = $result.UserPrincipalName
                UserDisplayName    = $result.UserDisplayName
                UserMailNickName   = $result.UserMailNickName
                Status             = $result.Status
                Message            = $result.Message
            }
        )
    }
}

$exportInfo = $results |
    Export-Results `
        -Path $OutputPath `
        -WorksheetName $OutputWorksheetName `
        -Force

$successCount = @(
    $results | Where-Object { $_.Status -in @('Added') }
).Count

$errorCount = @(
    $results | Where-Object { $_.Status -notin @('Added') }
).Count

Write-Host ("Operazioni completate. Riuscite: {0}. Errori: {1}. Report: {2}" -f $successCount, $errorCount, $exportInfo.Path)