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
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter()]
        [string] $InputPath,

        [Parameter()]
        [string] $OutputPath,

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $UseConsole,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $AllowClobber,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [switch] $Force
    )

    try {
        Write-Verbose "Connessione a Microsoft Entra..."
        Connect-ToEntra `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -ForceReconnect:$ForceReconnect `
            -AutoInstallModules:$AutoInstallModules `
            -AllowClobber:$AllowClobber `
            -Verbose:$VerbosePreference

        if (-not $InputPath) {
            $InputPath = Read-Host 'Percorso del file CSV di input (colonna "UPN")'
        }

        if ([string]::IsNullOrWhiteSpace($InputPath)) {
            throw "Percorso file di input non specificato."
        }

        if (-not (Test-Path $InputPath -PathType Leaf)) {
            throw "Il file di input '$InputPath' non esiste."
        }

        if (-not $OutputPath) {
            Write-Verbose "Richiesta percorso di esportazione..."
            $destination = Get-ExportDestination `
                -DefaultFileName 'UsersMailNicknames.xlsx' `
                -Formats @('xlsx','csv') `
                -PreferredFormat 'xlsx' `
                -Title 'Scegli dove salvare l''export dei MailNickName utenti' `
                -UseConsole:$UseConsole `
                -Force:$Force

            $OutputPath = $destination.Path
        }

        Write-Verbose "Lettura file CSV di input..."
        $inputRows = Import-Csv -Path $InputPath -ErrorAction Stop

        if (-not $inputRows) {
            throw "Il file CSV di input è vuoto."
        }

        if ('UPN' -notin $inputRows[0].PSObject.Properties.Name) {
            throw "Il file CSV deve contenere una colonna 'UPN'."
        }

        $results = foreach ($row in $inputRows) {
            $userUPN = $row.UPN

            if ([string]::IsNullOrWhiteSpace($userUPN)) {
                continue
            }

            $escapedUserUPN = $userUPN -replace "'", "''"
            Write-Verbose "Ricerca utente: $userUPN"

            $user = Get-EntraUser -Filter "UserPrincipalName eq '$escapedUserUPN'" -ErrorAction Stop

            if ($null -eq $user) {
                [pscustomobject]@{
                    ObjectId     = $null
                    UPN          = $userUPN
                    MailNickName = $null
                    Found        = $false
                }
                continue
            }

            [pscustomobject]@{
                ObjectId     = $user.ObjectId
                UPN          = $user.UserPrincipalName
                MailNickName = $user.MailNickName
                Found        = $true
            }
        }

        $finalResults = $results | Sort-Object UPN

        if (-not $finalResults) {
            Write-Warning "Nessun risultato da esportare."
            return
        }

        if ($PSCmdlet.ShouldProcess($OutputPath, "Esporta MailNickName utenti")) {
            Write-Verbose "Esportazione risultati in corso..."
            $export = Export-Results `
                -InputObject $finalResults `
                -Path $OutputPath `
                -WorksheetName 'UsersMailNicknames' `
                -Force:$Force

            Write-Verbose "Esportazione completata: $($export.Path)"
            return $export
        }
    }
    catch {
        $message = "Errore durante l'esportazione dei MailNickName utenti: $($_.Exception.Message)"
        Write-Error $message
    }
}