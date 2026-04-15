function Get-SipFederationTlsSrvRecords {
    <#
    .SYNOPSIS
        Recupera i record DNS SRV _sipfederationtls._tcp per i domini del tenant Microsoft 365.

    .DESCRIPTION
        Si connette a Microsoft Graph e risolve il record SRV _sipfederationtls._tcp
        per ciascun dominio. Se viene fornito -Domain, usa quella lista; altrimenti
        recupera tutti i domini del tenant tramite Get-MgDomain.

    .PARAMETER Domain
        Uno o più domini da interrogare. Accetta input da pipeline.
        Se omesso, vengono usati tutti i domini del tenant.

    .PARAMETER Scopes
        Scope Microsoft Graph da usare. Default: 'Domain.Read.All'.

    .PARAMETER TenantId
        Tenant ID da usare per la connessione a Microsoft Graph.

    .PARAMETER UseDeviceCode
        Usa l'autenticazione device code per la connessione a Microsoft Graph.

    .PARAMETER AutoInstallModules
        Installa automaticamente i moduli mancanti richiesti.

    .PARAMETER ForceReconnect
        Forza una nuova connessione a Microsoft Graph.

    .PARAMETER DnsServer
        Server DNS specifico per la risoluzione SRV. Se omesso usa il resolver di sistema.

    .OUTPUTS
        PSCustomObject per ogni record SRV trovato, con proprietà:
        Domain, QueryName, NameTarget, Port, Priority, Weight, Ttl,
        IsDefault, IsInitial, IsVerified, Authentication.

    .EXAMPLE
        Get-SipFederationTlsSrvRecords

    .EXAMPLE
        Get-SipFederationTlsSrvRecords -Domain "contoso.com", "fabrikam.com"

    .EXAMPLE
        "contoso.com", "fabrikam.com" | Get-SipFederationTlsSrvRecords -DnsServer "8.8.8.8"
    #>
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]] $Domain,

        [Parameter()]
        [string[]] $Scopes = @('Domain.ReadWrite.All'),

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [string] $DnsServer
    )

    begin {
        Write-Verbose "Connessione a Microsoft Graph..."
        Connect-ToGraph `
            -Scopes             $Scopes `
            -TenantId           $TenantId `
            -UseDeviceCode:     $UseDeviceCode `
            -AutoInstallModules:$AutoInstallModules `
            -ForceReconnect:    $ForceReconnect `
            -Verbose:           $VerbosePreference

        # Se non vengono forniti domini specifici, recupera tutti i domini del tenant
        $allTenantDomains = $null
        if (-not $Domain) {
            Write-Verbose "Recupero tutti i domini del tenant da Microsoft Graph..."
            $allTenantDomains = Get-MgDomain -ErrorAction Stop
        }

        $results = [System.Collections.Generic.List[object]]::new()
    }

    process {
        # Determina la lista su cui iterare in questo invocation
        $targets = if ($Domain) {
            # Costruisce oggetti minimali compatibili con il loop sottostante
            $Domain | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                ForEach-Object { [pscustomobject]@{
                    Id               = $_
                    IsDefault        = $null
                    IsInitial        = $null
                    IsVerified       = $null
                    AuthenticationType = $null
                } }
        } else {
            $allTenantDomains
        }

        foreach ($dom in $targets) {
            $domainName = $dom.Id
            if ([string]::IsNullOrWhiteSpace($domainName)) { continue }

            $recordName = "_sipfederationtls._tcp.$domainName"
            Write-Verbose "Risoluzione SRV per: $recordName"

            $dnsInfo = Resolve-SrvRecordSafe `
                -Name    $recordName `
                -Server  $DnsServer `
                -Verbose:$VerbosePreference

            if ($null -eq $dnsInfo) { continue }

            foreach ($rec in $dnsInfo) {
                $results.Add([pscustomobject]@{
                    Domain         = $domainName
                    QueryName      = $recordName
                    NameTarget     = $rec.NameTarget
                    Port           = $rec.Port
                    Priority       = $rec.Priority
                    Weight         = $rec.Weight
                    Ttl            = $rec.TTL
                    IsDefault      = $dom.IsDefault
                    IsInitial      = $dom.IsInitial
                    IsVerified     = $dom.IsVerified
                    Authentication = $dom.AuthenticationType
                })
            }
        }
    }

    end {
        Write-Verbose "Totale record trovati: $($results.Count)"

        if ($results.Count -eq 0) {
            Write-Warning "Nessun record SRV _sipfederationtls._tcp trovato."
            return
        }

        $results.ToArray() | Sort-Object Domain, Priority, Weight
    }
}