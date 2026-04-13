#Requires -Version 7.0

[CmdletBinding()]
param(
    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [string[]] $Scopes = @('Domain.ReadWrite.All'),

    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [switch] $UseDeviceCode,

    # Forza modalità console (no GUI)
    [Parameter()]
    [switch] $UseConsole,

    [Parameter()]
    [switch] $AutoInstallModules
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
    
# Import moduli locali (repo-friendly)
$modulesRoot = Join-Path $PSScriptRoot '..\Modules'
Import-Module (Join-Path $modulesRoot 'CloudConnect\CloudConnect.psd1') -Force
Import-Module (Join-Path $modulesRoot 'DnsTools\DnsTools.psd1') -Force
Import-Module (Join-Path $modulesRoot 'FilesUtilities\FilesUtilities.psd1') -Force

# 1) Connessione Graph
Connect-ToGraph -Scopes $Scopes -TenantId $TenantId -UseDeviceCode:$UseDeviceCode -AutoInstallModules:$AutoInstallModules

# 2) OutputPath (se non fornito)
if (-not $OutputPath) {
    $dest = Get-ExportDestination -DefaultFileName "sipfederationtls_srv_records.xlsx" `
                                  -Formats @("xlsx","csv") -PreferredFormat "xlsx" -UseConsole:$UseConsole
    $OutputPath = $dest.Path
}

# 3) Recupero domini
Write-Verbose "Recupero domini da Microsoft Graph..."
$domains = Get-MgDomain

# 4) Risultati
$results = New-Object System.Collections.Generic.List[object]

foreach ($domain in $domains) {
    $domainName = $domain.Id
    $recordName = "_sipfederationtls._tcp.$domainName"

    Write-Verbose "Risoluzione SRV per: $recordName"

    $dnsInfo = Resolve-SrvRecordSafe -Name $recordName
    if ($null -eq $dnsInfo) { continue }

    foreach ($rec in $dnsInfo) {
        $results.Add([pscustomobject]@{
            Domain     = $domainName
            QueryName  = $recordName
            NameTarget = $rec.NameTarget
            Port       = $rec.Port
            Priority   = $rec.Priority
            Weight     = $rec.Weight
            Ttl        = $rec.TTL
        })
    }
}

Write-Verbose "Totale record trovati: $($results.Count)"
$results | Sort-Object Domain, Priority, Weight | Format-Table -AutoSize

# 5) Export
$export = Export-Results -InputObject $results.ToArray() -Path $OutputPath -WorksheetName "SRV_sipfederationtls"

if ($export.Format -eq "xlsx") {
    Write-Host "✅ Export completato: $($export.Path) (worksheet: SRV_sipfederationtls)" -ForegroundColor Green
}
else {
    Write-Host "⚠️ Export in CSV: $($export.Path) (ImportExcel non installato o richiesto CSV)" -ForegroundColor Yellow
}