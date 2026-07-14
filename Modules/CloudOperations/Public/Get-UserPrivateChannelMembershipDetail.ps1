function Get-UserPrivateChannelMembershipDetail {
    <#
    .SYNOPSIS
    Recupera i canali privati di cui l'utente è membro per i Team specificati, includendo il ruolo.

    .DESCRIPTION
    Accetta in input le righe prodotte da Get-UserTeamsMembershipDetail e, per ogni Team,
    enumera i canali privati di cui l'utente è membro.

    Restituisce una riga per ogni relazione utente-Team-canale privato.
    I valori restituiti per ChannelRole sono:
    - Owner
    - Member

    .PARAMETER TeamMembership
    Oggetto di membership utente-Team prodotto da Get-UserTeamsMembershipDetail.
    Accetta input da pipeline.

    .PARAMETER TenantId
    Tenant ID da usare per la connessione a Microsoft Graph.

    .PARAMETER UseDeviceCode
    Usa l'autenticazione device code.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti.

    .PARAMETER ForceReconnect
    Forza una nuova connessione a Microsoft Graph.

    .PARAMETER IncludeTeamsWithoutPrivateChannels
    Se specificato, restituisce una riga anche per i Team in cui l'utente non appartiene
    ad alcun canale privato.

    .OUTPUTS
    PSCustomObject con:
    SourceIdentity, SourceIdentityType, UserObjectId, UserPrincipalName, UserDisplayName,
    UserMailNickName, TeamId, TeamDisplayName, TeamRole, ChannelId, ChannelDisplayName,
    ChannelMembershipType, ChannelRole.

    .EXAMPLE
    $teamRows | Get-UserPrivateChannelMembershipDetail

    .EXAMPLE
    $users |
        Get-UserTeamsMembershipDetail -TenantId $TenantId |
        Get-UserPrivateChannelMembershipDetail -TenantId $TenantId -IncludeTeamsWithoutPrivateChannels
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [Alias('InputObject')]
        [object[]] $TeamMembership,

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [switch] $IncludeTeamsWithoutPrivateChannels
    )

    begin {
        Write-Verbose "Connessione a Microsoft Graph..."
        Connect-ToGraph `
            -Scopes @('Group.Read.All', 'Channel.ReadBasic.All', 'ChannelMember.Read.All', 'User.Read.All') `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -ForceReconnect:$ForceReconnect `
            -AutoInstallModules:$AutoInstallModules `
            -Verbose:$VerbosePreference

        $requiredModules = @(
            'Microsoft.Graph.Teams'
        )

        foreach ($module in $requiredModules) {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                if ($AutoInstallModules) {
                    Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                }
                else {
                    throw "Modulo richiesto non trovato: $module"
                }
            }

            if (-not (Get-Module -Name $module)) {
                Import-Module $module -ErrorAction Stop -Verbose:$false | Out-Null
            }
        }
    }

    process {
        foreach ($row in $TeamMembership) {
            if ($null -eq $row) {
                continue
            }

            if ([string]::IsNullOrWhiteSpace($row.TeamId)) {
                Write-Warning "TeamId mancante nell'oggetto di input."
                continue
            }

            if ([string]::IsNullOrWhiteSpace($row.UserObjectId)) {
                Write-Warning "UserObjectId mancante per il Team '$($row.TeamDisplayName)'."
                continue
            }

            Write-Verbose "Recupero canali privati del Team '$($row.TeamDisplayName)' per '$($row.UserPrincipalName)'..."

            try {
                $channels = @(Get-MgTeamChannel -TeamId $row.TeamId -All -ErrorAction Stop |
                    Where-Object { $_.MembershipType -eq 'private' })
            }
            catch {
                Write-Warning "Impossibile leggere i canali del Team '$($row.TeamDisplayName)': $($_.Exception.Message)"
                continue
            }

            $outputCount = 0

            foreach ($channel in $channels) {
                try {
                    $channelMembers = @(Get-MgTeamChannelMember -TeamId $row.TeamId -ChannelId $channel.Id -All -ErrorAction Stop)

                    $channelMembership = $channelMembers |
                        Where-Object {
                            $_.UserId -eq $row.UserObjectId
                        } |
                        Select-Object -First 1

                    if ($null -eq $channelMembership) {
                        continue
                    }

                    $channelRole = if ($channelMembership.Roles -contains 'owner') { 'Owner' } else { 'Member' }

                    $outputCount++

                    [pscustomobject][ordered]@{
                        SourceIdentity         = $row.SourceIdentity
                        SourceIdentityType     = $row.SourceIdentityType
                        UserObjectId           = $row.UserObjectId
                        UserPrincipalName      = $row.UserPrincipalName
                        UserDisplayName        = $row.UserDisplayName
                        UserMailNickName       = $row.UserMailNickName
                        TeamId                 = $row.TeamId
                        TeamDisplayName        = $row.TeamDisplayName
                        TeamRole               = $row.TeamRole
                        ChannelId              = $channel.Id
                        ChannelDisplayName     = $channel.DisplayName
                        ChannelMembershipType  = $channel.MembershipType
                        ChannelRole            = $channelRole
                    }
                }
                catch {
                    Write-Warning "Errore nel canale '$($channel.DisplayName)' del Team '$($row.TeamDisplayName)': $($_.Exception.Message)"
                }
            }

            if ($IncludeTeamsWithoutPrivateChannels -and $outputCount -eq 0) {
                [pscustomobject][ordered]@{
                    SourceIdentity         = $row.SourceIdentity
                    SourceIdentityType     = $row.SourceIdentityType
                    UserObjectId           = $row.UserObjectId
                    UserPrincipalName      = $row.UserPrincipalName
                    UserDisplayName        = $row.UserDisplayName
                    UserMailNickName       = $row.UserMailNickName
                    TeamId                 = $row.TeamId
                    TeamDisplayName        = $row.TeamDisplayName
                    TeamRole               = $row.TeamRole
                    ChannelId              = $null
                    ChannelDisplayName     = $null
                    ChannelMembershipType  = $null
                    ChannelRole            = $null
                }
            }
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}