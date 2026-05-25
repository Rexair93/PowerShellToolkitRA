function Get-EntraAppSecrets {
    <#
    .SYNOPSIS
    Cerca le app registrate in Entra ID il cui nome contiene una stringa
    e restituisce i dati dei secret associati, con logica ottimizzata per tenant grandi.

    .DESCRIPTION
    Si connette a Microsoft Graph e usa una ricerca indicizzata lato Graph
    tramite -Search su DisplayName per limitare il numero di application
    recuperate. Successivamente rifinisce localmente i risultati per ottenere
    un comportamento più vicino a un "contains" sul DisplayName.

    Se un'app non ha secret, viene comunque restituita una riga con
    SecretEndDateTime e SecretId valorizzati a $null.

    .PARAMETER SearchString
    Stringa da cercare nel DisplayName delle app registrate.

    .PARAMETER Scopes
    Scope Microsoft Graph da usare. Default: Application.Read.All.

    .PARAMETER TenantId
    Tenant ID da usare per la connessione a Microsoft Graph.

    .PARAMETER UseDeviceCode
    Usa l'autenticazione device code per la connessione a Microsoft Graph.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti richiesti.

    .PARAMETER ForceReconnect
    Forza una nuova connessione a Microsoft Graph.

    .PARAMETER PageSize
    Numero di elementi per pagina richiesti a Graph.

    .PARAMETER MaxResults
    Numero massimo di application da elaborare dopo la ricerca Graph.
    0 = nessun limite.

    .OUTPUTS
    PSCustomObject con DisplayName, ApplicationClientId, SecretEndDateTime, SecretId, HasSecret.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('Name', 'DisplayName', 'Filter', 'Query')]
        [string[]] $SearchString,

        [Parameter()]
        [string[]] $Scopes = @('Application.Read.All'),

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [ValidateRange(1, 999)]
        [int] $PageSize = 100,

        [Parameter()]
        [ValidateRange(0, 1000000)]
        [int] $MaxResults = 1000
    )

    begin {
        Write-Verbose "Connessione a Microsoft Graph..."
        Connect-ToGraph `
            -Scopes $Scopes `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -ForceReconnect:$ForceReconnect `
            -AutoInstallModules:$AutoInstallModules `
            -Verbose:$VerbosePreference
    }

    process {
        foreach ($term in $SearchString) {
            if ([string]::IsNullOrWhiteSpace($term)) {
                continue
            }

            $currentTerm = $term.Trim()
            $escapedSearchTerm = $currentTerm.Replace('"', '\"')
            $escapedLikeTerm = $currentTerm.Replace('[', '[[]').Replace('*', '[*]').Replace('?', '[?]')

            Write-Verbose "Ricerca indicizzata application con DisplayName contenente: $currentTerm"

            try {
                $apps = Get-MgApplication `
                    -Search "`"displayName:$escapedSearchTerm`"" `
                    -ConsistencyLevel eventual `
                    -Property @('Id', 'AppId', 'DisplayName', 'PasswordCredentials') `
                    -PageSize $PageSize `
                    -All `
                    -ErrorAction Stop
            }
            catch {
                throw "Errore durante la ricerca delle application con '$currentTerm': $($_.Exception.Message)"
            }

            if ($null -eq $apps) {
                continue
            }

            $filteredApps = $apps |
                Where-Object {
                    $_.DisplayName -and $_.DisplayName -like "*$escapedLikeTerm*"
                } |
                Sort-Object DisplayName, AppId

            if ($MaxResults -gt 0) {
                $filteredApps = $filteredApps | Select-Object -First $MaxResults
            }

            foreach ($app in $filteredApps) {
                if ($app.PasswordCredentials -and $app.PasswordCredentials.Count -gt 0) {
                    foreach ($secret in $app.PasswordCredentials) {
                        [PSCustomObject]@{
                            SearchString         = $currentTerm
                            DisplayName          = $app.DisplayName
                            ApplicationClientId  = $app.AppId
                            ObjectID             = $app.Id
                            SecretStartDateTime  = $secret.StartDateTime
                            SecretEndDateTime    = $secret.EndDateTime
                            SecretId             = $secret.KeyId
                            HasSecret            = $true
                        }
                    }
                }
                else {
                    [PSCustomObject]@{
                        SearchString         = $currentTerm
                        DisplayName          = $app.DisplayName
                        ApplicationClientId  = $app.AppId
                        ObjectID             = $app.Id
                        SecretStartDateTime  = $null
                        SecretEndDateTime    = $null
                        SecretId             = $null
                        HasSecret            = $false
                    }
                }
            }
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}