function Get-EntraUserProperties {
    <#
    .SYNOPSIS
    Recupera proprietà arbitrarie di utenti Entra ID a partire da un parametro di ricerca configurabile.

    .DESCRIPTION
    Accetta una lista di valori (es. EmployeeId, UPN, mail...), li usa per filtrare utenti
    tramite Get-EntraUser e restituisce le proprietà richieste. Se il parametro di lookup o
    una proprietà richiesta non esiste, emette un avviso con i valori validi disponibili.

    .PARAMETER InputValues
    Lista di valori su cui eseguire la ricerca (es. lista di EmployeeId, UPN, mail...).

    .PARAMETER LookupProperty
    Proprietà Entra su cui filtrare la ricerca. Default: 'UserPrincipalName'.
    Se il valore specificato non esiste come proprietà filtrabile, viene emesso un avviso
    con le proprietà comunemente utilizzabili.

    .PARAMETER Properties
    Elenco delle proprietà da includere nell'output. Default: 'ObjectId','UserPrincipalName','MailNickName'.
    Le proprietà non trovate sull'oggetto utente vengono segnalate con un avviso e sostituite con $null.

    .PARAMETER TenantId
    Tenant ID per la connessione a Microsoft Entra.

    .PARAMETER UseDeviceCode
    Usa l'autenticazione device code.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti.

    .PARAMETER AllowClobber
    Consente AllowClobber durante l'installazione del modulo Microsoft.Entra.

    .PARAMETER ForceReconnect
    Forza una nuova connessione a Microsoft Entra.

    .OUTPUTS
    PSCustomObject contenente LookupValue, Found e le proprietà richieste.

    .EXAMPLE
    # Recupera UPN, ObjectId e MailNickName a partire da una lista di EmployeeId
    Get-EntraUserProperties -InputValues '12345','67890' -LookupProperty 'EmployeeId' -Properties 'UserPrincipalName','ObjectId','MailNickName'

    .EXAMPLE
    # Usa i default: ricerca per UPN, restituisce ObjectId, UserPrincipalName, MailNickName
    Get-EntraUserProperties -InputValues 'mario.rossi@contoso.com','luigi.bianchi@contoso.com'

    .EXAMPLE
    # Pipeline da CSV
    Import-Csv .\employees.csv | Select-Object -ExpandProperty EmployeeId |
        Get-EntraUserProperties -LookupProperty 'EmployeeId'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]] $InputValues,

        [Parameter()]
        [string] $LookupProperty = 'UserPrincipalName',

        [Parameter()]
        [string[]] $Properties = @('ObjectId', 'UserPrincipalName', 'GivenName', 'Surname', 'DisplayName', 'MailNickName', 'EmployeeId'),

        [Parameter()]
        [string[]] $Scopes = @('User.Read.All'),

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $AllowClobber,

        [Parameter()]
        [switch] $ForceReconnect
    )

    begin {
        # Proprietà Entra ID comunemente utilizzabili come filtro OData
        $script:KnownFilterableProperties = @(
            'AccountEnabled', 'AgeGroup', 'AssignedLicenses', 'City', 'CompanyName',
            'ConsentProvidedForMinor', 'Country', 'CreatedDateTime', 'Department',
            'DisplayName', 'EmployeeId', 'GivenName', 'JobTitle', 'Mail',
            'MailNickName', 'MobilePhone', 'ObjectId', 'OnPremisesImmutableId',
            'OnPremisesSecurityIdentifier', 'OnPremisesSamAccountName',
            'OnPremisesUserPrincipalName', 'PostalCode', 'PreferredLanguage',
            'State', 'StreetAddress', 'Surname', 'UsageLocation',
            'UserPrincipalName', 'UserType'
        )

        Write-Verbose "Connessione a Microsoft Entra..."
        Connect-ToEntra `
            -Scopes $Scopes `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -ForceReconnect:$ForceReconnect `
            -AutoInstallModules:$AutoInstallModules `
            -AllowClobber:$AllowClobber `
            -Verbose:$VerbosePreference

        # Valida LookupProperty una sola volta nel begin
        if ($LookupProperty -notin $script:KnownFilterableProperties) {
            $suggestion = ($script:KnownFilterableProperties | Sort-Object) -join ', '
            Write-Warning (
                "La proprietà di lookup '$LookupProperty' non è tra quelle note come filtrabili su Entra ID. " +
                "Potrebbe comunque funzionare se supportata dall'API, ma in caso di errore considera una di queste: $suggestion"
            )
        }

        $script:PropertiesValidated = $false
    }

    process {
        foreach ($value in $InputValues) {
            if ([string]::IsNullOrWhiteSpace($value)) {
                continue
            }

            $escapedValue = $value -replace "'", "''"
            Write-Verbose "Ricerca utente con $LookupProperty = '$value'"

            try {
                $user = Get-EntraUser -Filter "$LookupProperty eq '$escapedValue'" -ErrorAction Stop
            }
            catch {
                # Intercetta errori tipici di proprietà non filtrabili
                $errMsg = $_.Exception.Message
                if ($errMsg -match 'Invalid filter|unsupported|not supported|filterable') {
                    $suggestion = ($script:KnownFilterableProperties | Sort-Object) -join ', '
                    Write-Warning (
                        "La proprietà '$LookupProperty' non sembra supportare il filtraggio OData su Entra ID. " +
                        "Proprietà note come filtrabili: $suggestion"
                    )
                }
                throw "Errore durante la ricerca con $LookupProperty='$value': $errMsg"
            }

            if ($null -eq $user) {
                # Costruisce l'oggetto con tutte le Properties a $null
                $output = [ordered]@{ LookupValue = $value; Found = $false }
                foreach ($prop in $Properties) {
                    $output[$prop] = $null
                }
                [pscustomobject]$output
                continue
            }

            # Al primo utente trovato, valida le Properties richieste
            if (-not $script:PropertiesValidated) {
                $availableProps = $user.PSObject.Properties.Name
                foreach ($prop in $Properties) {
                    if ($prop -notin $availableProps) {
                        Write-Warning (
                            "La proprietà '$prop' non è presente nell'oggetto utente restituito da Entra ID. " +
                            "Proprietà disponibili: $($availableProps -join ', ')"
                        )
                    }
                }
                $script:PropertiesValidated = $true
            }

            $output = [ordered]@{ LookupValue = $value; Found = $true }
            foreach ($prop in $Properties) {
                $output[$prop] = $user.$prop
            }
            [pscustomobject]$output
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}
