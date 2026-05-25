[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$TeamsAppExternalId,

    [Parameter()]
    [ValidateSet('xlsx','csv')]
    [string]$Format = 'xlsx',

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [ValidateRange(1,20)]
    [int]$BatchSize = 20,

    [Parameter()]
    [switch]$UseDeviceCode,

    [Parameter()]
    [switch]$IncludeAllUsers,

    [Parameter()]
    [string]$UserFilter = "accountEnabled eq true and userType eq 'Member'",

    [Parameter()]
    [switch]$AutoInstallModules,

    [Parameter()]
    [switch]$ForceReconnect
)

Import-Module FilesUtilities -ErrorAction SilentlyContinue
Import-Module CloudOperations -ErrorAction SilentlyContinue

if ($AutoInstallModules) {
    Assert-Module -Name Microsoft.Graph -Scope CurrentUser -AutoInstall
}

if (-not $OutputPath) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputPath = Join-Path (Get-Location) "TeamsAppInstallationReport_$timestamp.$Format"
}

$results = Get-TeamsAppInstallationReport `
    -TeamsAppExternalId $TeamsAppExternalId `
    -BatchSize $BatchSize `
    -UseDeviceCode:$UseDeviceCode `
    -IncludeAllUsers:$IncludeAllUsers `
    -UserFilter $UserFilter `
    -AutoInstallModules:$AutoInstallModules `
    -ForceReconnect:$ForceReconnect `
    -Verbose:$VerbosePreference

$export = Export-Results -InputObject $results -Path $OutputPath -WorksheetName 'TeamsAppInstallations' -Force

[pscustomobject]@{
    Path               = $export.Path
    Format             = $export.Format
    TotalUsers         = @($results).Count
    InstalledUsers     = @($results | Where-Object IsInstalled).Count
    NotInstalledUsers  = @($results | Where-Object { -not $_.IsInstalled }).Count
}