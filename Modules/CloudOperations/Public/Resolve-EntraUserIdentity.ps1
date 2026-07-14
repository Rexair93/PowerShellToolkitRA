function Resolve-EntraUserIdentity {
    <#
    .SYNOPSIS
    Risolve un'identità utente in un profilo Entra normalizzato.

    .DESCRIPTION
    Accetta UPN, ObjectId o MailNickName e restituisce sempre un oggetto canonico
    con ObjectId, UserPrincipalName, MailNickName e DisplayName.

    .PARAMETER Identity
    Valore da risolvere.

    .PARAMETER IdentityType
    Tipo di identità in input: Auto, UserPrincipalName, ObjectId, MailNickName.

    .PARAMETER TenantId
    Tenant ID da usare per la connessione a Entra.

    .PARAMETER UseDeviceCode
    Usa autenticazione device code.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti.

    .PARAMETER AllowClobber
    Consente AllowClobber durante installazione moduli.

    .PARAMETER ForceReconnect
    Forza una nuova connessione.

    .OUTPUTS
    PSCustomObject
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string[]] $Identity,

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
        [switch] $ForceReconnect
    )

    process {
        foreach ($value in $Identity) {
            if ([string]::IsNullOrWhiteSpace($value)) {
                continue
            }

            $value = $value.Trim()

            $effectiveIdentityType = $IdentityType
            if ($effectiveIdentityType -eq 'Auto') {
                $effectiveIdentityType = if ($value -match '@') {
                    'UserPrincipalName'
                }
                elseif ([guid]::TryParse($value, [ref]([guid]::Empty))) {
                    'ObjectId'
                }
                else {
                    'MailNickName'
                }
            }

            try {
                $resolved = @(Get-EntraUserProperties `
                    -InputValues         $value `
                    -LookupProperty      $effectiveIdentityType `
                    -Properties          'ObjectId', 'UserPrincipalName', 'MailNickName', 'DisplayName' `
                    -TenantId            $TenantId `
                    -UseDeviceCode:      $UseDeviceCode `
                    -AutoInstallModules: $AutoInstallModules `
                    -AllowClobber:       $AllowClobber `
                    -ForceReconnect:     $ForceReconnect `
                    -Verbose:            $VerbosePreference)

                $match = $resolved | Where-Object { $_.Found } | Select-Object -First 1

                if ($null -eq $match) {
                    [pscustomobject][ordered]@{
                        SourceIdentity       = $value
                        SourceIdentityType   = $effectiveIdentityType
                        Found                = $false
                        ObjectId             = $null
                        UserPrincipalName    = $null
                        MailNickName         = $null
                        DisplayName          = $null
                    }
                    continue
                }

                [pscustomobject][ordered]@{
                    SourceIdentity       = $value
                    SourceIdentityType   = $effectiveIdentityType
                    Found                = $true
                    ObjectId             = $match.ObjectId
                    UserPrincipalName    = $match.UserPrincipalName
                    MailNickName         = $match.MailNickName
                    DisplayName          = $match.DisplayName
                }
            }
            catch {
                Write-Warning "Errore nella risoluzione dell'identità '$value': $($_.Exception.Message)"
            }
        }
    }
}