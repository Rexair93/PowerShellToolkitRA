function Get-SipFederationTlsSrvRecords {
    <#
    .SYNOPSIS
    Recupera ed esporta i record DNS SRV _sipfederationtls._tcp dei domini del tenant Microsoft 365.

    .DESCRIPTION
    Si connette a Microsoft Graph, legge i domini del tenant, risolve per ciascuno
    il record SRV _sipfederationtls._tcp e ne esporta i risultati in formato XLSX
    o CSV tramite le utility condivise del modulo.

    .PARAMETER OutputPath
    Percorso completo del file di output. Se non specificato, viene richiesto tramite
    Get-ExportDestination.

    .PARAMETER Scopes
    Scope Microsoft Graph da usare per il recupero dei domini.

    .PARAMETER TenantId
    Tenant ID da usare per la connessione a Microsoft Graph.

    .PARAMETER UseDeviceCode
    Usa l'autenticazione device code per la connessione a Microsoft Graph.

    .PARAMETER UseConsole
    Forza la scelta del file di destinazione in modalità console.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti richiesti.

    .PARAMETER DnsServer
    Server DNS specifico da usare per la risoluzione dei record SRV.

    .PARAMETER Force
    Consente la sovrascrittura del file di output quando supportato.

    .OUTPUTS
    PSCustomObject con Path e Format.

    .EXAMPLE
    Export-SipFederationTlsSrvRecords

    .EXAMPLE
    Export-SipFederationTlsSrvRecords -UseConsole -Verbose

    .EXAMPLE
    Export-SipFederationTlsSrvRecords -TenantId "contoso.onmicrosoft.com" -UseDeviceCode
    #>

    [CmdletBinding(SupportsShouldProcess)]
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
        [switch] $AutoInstallModules,

        [Parameter()]
        [string] $DnsServer,

        [Parameter()]
        [switch] $Force
    )

    try {
    # 1) Connessione Graph
    Write-Verbose "Connessione a Microsoft Graph..."
    Connect-ToGraph `
        -Scopes $Scopes `
        -TenantId $TenantId `
        -UseDeviceCode:$UseDeviceCode `
        -AutoInstallModules:$AutoInstallModules `
        -Verbose:$VerbosePreference

    # 2) OutputPath (se non fornito)
    if (-not $OutputPath) {
        Write-Verbose "Richiesta percorso di esportazione..."
        $destination = Get-ExportDestination `
            -DefaultFileName "sipfederationtls_srv_records.xlsx" `
            -Formats @("xlsx","csv") `
            -PreferredFormat "xlsx" `
            -Title 'Scegli dove salvare l''export dei record SRV sipfederationtls' `
            -UseConsole:$UseConsole `
            -Force:$Force
        
        $OutputPath = $destination.Path
    }

    # 3) Recupero domini
    Write-Verbose "Recupero domini da Microsoft Graph..."
    $domains = Get-MgDomain -ErrorAction Stop

    # 4) Risultati
    $results = New-Object System.Collections.Generic.List[object]

    foreach ($domain in $domains) {
        $domainName = $domain.Id

        if ([string]::IsNullOrWhiteSpace($domainName)) {
            continue
        }
        
        $recordName = "_sipfederationtls._tcp.$domainName"
        Write-Verbose "Risoluzione SRV per: $recordName"

        $dnsInfo = Resolve-SrvRecordSafe -Name $recordName -Server $DnsServer -Verbose:$VerbosePreference
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
                IsDefault     = $domain.IsDefault
                IsInitial     = $domain.IsInitial
                IsVerified    = $domain.IsVerified
                Authentication = $domain.AuthenticationType
            })
        }
    }

    Write-Verbose "Totale record trovati: $($results.Count)"
    $finalResults = $results.ToArray() | Sort-Object Domain, Priority, Weight

    if (-not $finalResults) {
        Write-Warning "Nessun record SRV _sipfederationtls._tcp trovato per i domini del tenant."
        return
    }

    # 5) Export
    if ($PSCmdlet.ShouldProcess($OutputPath, "Esporta record SRV sipfederationtls")) {
        Write-Verbose "Esportazione risultati in corso..."
        $export = Export-Results `
            -InputObject $finalResults `
            -Path $OutputPath `
            -WorksheetName 'SRV_sipfederationtls' `
            -Force:$Force

        Write-Verbose "Esportazione completata: $($export.Path)"
        return $export
    }
    }
    catch {
            $message = "Errore durante l'esportazione dei record SRV sipfederationtls: $($_.Exception.Message)"
            Write-Error $message
        }
}