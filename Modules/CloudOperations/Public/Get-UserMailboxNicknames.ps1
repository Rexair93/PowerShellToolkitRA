function Get-UserMailboxNicknames {
    <#
    .SYNOPSIS
    Recupera il MailNickName degli utenti specificati in un file CSV ed esporta i risultati.

    .DESCRIPTION
    Legge un file CSV contenente una colonna 'UPN', si connette a Microsoft Entra,
    recupera per ogni utente ObjectId, UserPrincipalName e MailNickName, quindi
    esporta il risultato in formato XLSX o CSV.

    .PARAMETER InputPath
    Percorso del file CSV di input contenente la colonna 'UPN'.

    .PARAMETER OutputPath
    Percorso completo del file di output. Se non specificato, viene richiesto tramite
    Get-ExportDestination.

    .PARAMETER TenantId
    Tenant ID da usare per la connessione a Microsoft Entra.

    .PARAMETER UseDeviceCode
    Prova a usare l'autenticazione device code, se supportata dal modulo Microsoft.Entra.

    .PARAMETER UseConsole
    Forza la selezione dei percorsi in modalità console.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti richiesti.

    .PARAMETER AllowClobber
    Consente l'uso di AllowClobber durante l'installazione del modulo Microsoft.Entra.

    .PARAMETER ForceReconnect
    Forza una nuova connessione a Microsoft Entra.

    .PARAMETER Force
    Consente la sovrascrittura del file di output quando supportato.

    .OUTPUTS
    PSCustomObject con Path e Format.

    .EXAMPLE
    Export-UserMailboxNicknames -InputPath .\users.csv

    .EXAMPLE
    Export-UserMailboxNicknames -InputPath .\users.csv -UseConsole -Verbose
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('UPN')]
        [string[]] $UserPrincipalName,

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
        Write-Verbose "Connessione a Microsoft Entra..."
        Connect-ToEntra `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -ForceReconnect:$ForceReconnect `
            -AutoInstallModules:$AutoInstallModules `
            -AllowClobber:$AllowClobber `
            -Verbose:$VerbosePreference
    }

    process {
        foreach ($upn in $UserPrincipalName) {
            if ([string]::IsNullOrWhiteSpace($upn)) {
                continue
            }

            $escapedUpn = $upn -replace "'", "''"
            Write-Verbose "Ricerca utente: $upn"

            try {
                $user = Get-EntraUser -Filter "UserPrincipalName eq '$escapedUpn'" -ErrorAction Stop
            }
            catch {
                throw "Errore durante la ricerca dell'utente '$upn': $($_.Exception.Message)"
            }

            if ($null -eq $user) {
                [pscustomobject]@{
                    ObjectId          = $null
                    UserPrincipalName = $upn
                    MailNickName      = $null
                    Found             = $false
                }
                continue
            }

            [pscustomobject]@{
                ObjectId          = $user.ObjectId
                UserPrincipalName = $user.UserPrincipalName
                MailNickName      = $user.MailNickName
                Found             = $true
            }
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}