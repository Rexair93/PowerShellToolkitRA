function Get-TeamsAllowedDomains {
    <#
    .SYNOPSIS
    Recupera i domini consentiti della federazione Microsoft Teams.

    .DESCRIPTION
    Si connette a Microsoft Teams usando le utility del modulo, legge i domini
    consentiti da Get-CsTenantFederationConfiguration e li esporta tramite
    Export-Results. La destinazione viene scelta tramite GUI o console usando
    Get-ExportDestination.

    .PARAMETER TenantId
    Tenant ID da usare per la connessione a Microsoft Teams.

    .PARAMETER UseDeviceCode
    Prova a usare l'autenticazione device code, se supportata dal modulo MicrosoftTeams.

    .PARAMETER ForceReconnect
    Forza la disconnessione e una nuova connessione a Microsoft Teams.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti richiesti.

    .PARAMETER UseConsole
    Forza la selezione del file di destinazione in modalità console.

    .PARAMETER InitialDirectory
    Cartella iniziale proposta per il salvataggio.

    .PARAMETER DefaultFileName
    Nome file predefinito proposto in fase di esportazione.

    .PARAMETER Force
    Consente la sovrascrittura del file di output quando supportato.

    .OUTPUTS
    PSCustomObject con Path e Format.

    .EXAMPLE
    Export-TeamsAllowedDomains

    .EXAMPLE
    Export-TeamsAllowedDomains -UseConsole -Verbose

    .EXAMPLE
    Export-TeamsAllowedDomains -TenantId "contoso.onmicrosoft.com" -UseDeviceCode
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $UseConsole,

        [Parameter()]
        [string] $InitialDirectory = (Get-Location).Path,

        [Parameter()]
        [string] $DefaultFileName = 'Teams-AllowedDomains.xlsx',

        [Parameter()]
        [switch] $Force
    )

    try {
        Write-Verbose "Connessione a Microsoft Teams..."
        Connect-ToMicrosoftTeams `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -ForceReconnect:$ForceReconnect `
            -AutoInstallModules:$AutoInstallModules `
            -Verbose:$VerbosePreference

        Write-Verbose "Recupero configurazione domini consentiti da Teams..."
        $domains = Get-CsTenantFederationConfiguration -ErrorAction Stop |
            Select-Object -ExpandProperty AllowedDomains -ErrorAction Stop |
            Select-Object -ExpandProperty AllowedDomain -ErrorAction Stop |
            Select-Object -ExpandProperty Domain |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique |
            ForEach-Object {
                [pscustomobject]@{
                    Domain = $_
                }
            }

        if (-not $domains) {
            Write-Warning "Nessun dominio consentito trovato nella configurazione Teams."
            return
        }

        Write-Verbose "Richiesta percorso di esportazione..."
        $destination = Get-ExportDestination `
            -DefaultFileName $DefaultFileName `
            -InitialDirectory $InitialDirectory `
            -Formats @('xlsx', 'csv') `
            -PreferredFormat 'xlsx' `
            -Title "Scegli dove salvare l'export dei domini consentiti di Teams" `
            -UseConsole:$UseConsole `
            -Force:$Force

        if ($PSCmdlet.ShouldProcess($destination.Path, "Esporta domini consentiti Teams")) {
            Write-Verbose "Esportazione risultati in corso..."
            $result = Export-Results `
                -InputObject $domains `
                -Path $destination.Path `
                -WorksheetName 'AllowedDomains' `
                -Force:$Force

            Write-Verbose "Esportazione completata: $($result.Path)"
            return $result
        }
    }
    catch {
        $message = "Errore durante l'esportazione dei domini consentiti di Teams: $($_.Exception.Message)"
        Write-Error $message
    }
}