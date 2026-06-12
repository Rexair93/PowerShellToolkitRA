function Get-UserTeamMemberships {
    <#
    .SYNOPSIS
        Recupera i team Microsoft Teams di appartenenza per gli utenti specificati,
        restituendo ogni team una sola volta.

    .DESCRIPTION
        Per ogni utente della lista, recupera direttamente i team di cui è membro
        tramite Get-Team -User, senza caricare tutti i team del tenant.
        Ogni team viene aggiunto a un HashSet durante il processo: i duplicati
        vengono scartati in tempo reale, senza post-elaborazione finale.

        Se l'identificativo è già un UserPrincipalName, viene usato direttamente.
        Se non lo è, viene risolto tramite Entra ID usando la proprietà MailNickName,
        delegando connessione e lookup a Get-EntraUserProperties.

    .PARAMETER Identity
        Elenco di identità utente da cercare. Accetta input da pipeline.
        Il valore può essere un UserPrincipalName oppure un MailNickName.

    .PARAMETER IdentityType
        Tipo di identificativo usato nel parametro Identity.
        Valori ammessi:
        - Auto: se il valore contiene '@' viene trattato come UPN, altrimenti come MailNickName
        - UserPrincipalName: interpreta ogni valore come UPN
        - MailNickName: interpreta ogni valore come MailNickName

    .PARAMETER TenantId
        Tenant ID da usare per la connessione a Microsoft Teams e Microsoft Entra.

    .PARAMETER UseDeviceCode
        Prova a usare l'autenticazione device code, se supportata dai moduli in uso.

    .PARAMETER AutoInstallModules
        Installa automaticamente i moduli mancanti richiesti.

    .PARAMETER AllowClobber
        Consente l'uso di AllowClobber durante l'installazione del modulo Microsoft.Entra.

    .PARAMETER ForceReconnect
        Forza una nuova connessione ai servizi richiesti.

    .OUTPUTS
        PSCustomObject con proprietà: TeamId, TeamDisplayName.
        Ogni team è restituito una sola volta, indipendentemente da quanti
        utenti della lista ne siano membri.

    .EXAMPLE
        Get-UserTeamMemberships -Identity "mario.rossi@contoso.com"

    .EXAMPLE
        "mario.rossi@contoso.com", "lbianchi" | Get-UserTeamMemberships -IdentityType Auto -Verbose
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('UPN')]
        [string[]] $Identity,

        [Parameter()]
        [ValidateSet('Auto', 'UserPrincipalName', 'MailNickName')]
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

    begin {
        Write-Verbose "Connessione a Microsoft Teams..."
        Connect-ToMicrosoftTeams `
            -TenantId            $TenantId `
            -UseDeviceCode:      $UseDeviceCode `
            -ForceReconnect:     $ForceReconnect `
            -AutoInstallModules: $AutoInstallModules `
            -ImportMode          All `
            -Verbose:            $VerbosePreference

        # HashSet per la deduplicazione in tempo reale tramite TeamId
        $script:SeenTeamIds = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
    }

    process {
        foreach ($value in $Identity) {
            if ([string]::IsNullOrWhiteSpace($value)) {
                continue
            }

            $effectiveIdentityType = $IdentityType
            if ($IdentityType -eq 'Auto') {
                $effectiveIdentityType = if ($value -match '@') { 'UserPrincipalName' } else { 'MailNickName' }
            }

            $resolvedUpn = $null

            if ($effectiveIdentityType -eq 'UserPrincipalName') {
                Write-Verbose "Uso diretto dell'UPN '$value'."
                $resolvedUpn = $value
            }
            else {
                Write-Verbose "Risoluzione UPN da MailNickName '$value' tramite Entra ID..."

                try {
                    $resolvedUsers = @(
                        Get-EntraUserProperties `
                            -InputValues         $value `
                            -LookupProperty      'MailNickName' `
                            -Properties          'UserPrincipalName', 'MailNickName', 'DisplayName' `
                            -TenantId            $TenantId `
                            -UseDeviceCode:      $UseDeviceCode `
                            -AutoInstallModules: $AutoInstallModules `
                            -AllowClobber:       $AllowClobber `
                            -ForceReconnect:     $ForceReconnect `
                            -Verbose:            $VerbosePreference |
                        Where-Object { $_.Found -and -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) }
                    )
                }
                catch {
                    Write-Warning "Errore nella risoluzione Entra per '$value': $($_.Exception.Message)"
                    $resolvedUsers = @()
                }

                if ($resolvedUsers.Count -gt 1) {
                    Write-Warning "Trovati $($resolvedUsers.Count) utenti per MailNickName '$value'. Verrà usato il primo risultato."
                }

                $resolvedUpn = ($resolvedUsers | Select-Object -First 1).UserPrincipalName
            }

            if ([string]::IsNullOrWhiteSpace($resolvedUpn)) {
                Write-Warning "Impossibile risolvere un UserPrincipalName valido per '$value'. Utente ignorato."
                continue
            }

            Write-Verbose "Recupero team per l'utente '$resolvedUpn'..."

            try {
                $userTeams = Get-Team -User $resolvedUpn -ErrorAction Stop
            }
            catch {
                Write-Warning "Impossibile recuperare i team per '$resolvedUpn': $($_.Exception.Message)"
                continue
            }

            if (-not $userTeams) {
                Write-Verbose "Nessun team trovato per '$resolvedUpn'."
                continue
            }

            foreach ($team in $userTeams) {
                # Aggiunge solo se non già visto: Add restituisce $false se già presente
                if ($script:SeenTeamIds.Add($team.GroupId)) {
                    [pscustomobject]@{
                        TeamId          = $team.GroupId
                        TeamDisplayName = $team.DisplayName
                    }
                }
            }
        }
    }

    end {
        Write-Verbose "Elaborazione completata. Team univoci trovati: $($script:SeenTeamIds.Count)."
    }
}