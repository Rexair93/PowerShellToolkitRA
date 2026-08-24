function Get-UserTeamsMembershipDetail {
    <#
    .SYNOPSIS
    Recupera i Team Microsoft Teams di cui un utente è membro, includendo il ruolo nel Team.

    .DESCRIPTION
    Accetta uno o più utenti già risolti oppure identità utente da risolvere a monte
    e restituisce una riga per ogni relazione utente-Team.

    Il ruolo nel Team viene determinato interrogando la membership del Team via Microsoft Graph.
    I valori restituiti per TeamRole sono:
    - Owner
    - Member

    .PARAMETER User
    Oggetto utente canonico contenente almeno:
    ObjectId, UserPrincipalName, DisplayName, MailNickName.
    Accetta input da pipeline.

    .PARAMETER TenantId
    Tenant ID da usare per la connessione a Microsoft Graph.

    .PARAMETER UseDeviceCode
    Usa l'autenticazione device code.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti.

    .PARAMETER ForceReconnect
    Forza una nuova connessione a Microsoft Graph.

    .OUTPUTS
    PSCustomObject con:
    SourceIdentity, SourceIdentityType, UserObjectId, UserPrincipalName, UserDisplayName,
    UserMailNickName, TeamId, TeamDisplayName, TeamRole.

    .EXAMPLE
    $user | Get-UserTeamsMembershipDetail

    .EXAMPLE
    $users | Get-UserTeamsMembershipDetail -TenantId $TenantId -Verbose
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [object[]] $User,

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [ValidateSet('Microsoft.Graph', 'MicrosoftTeams')]
        [string] $SelectedModule = 'Microsoft.Graph',

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $ForceReconnect
    )

    begin {
        switch ($SelectedModule) {
            'Microsoft.Graph' {
                Write-Verbose "Connessione a Microsoft Graph..."
                Connect-ToGraph `
                    -Scopes @('Group.Read.All', 'User.Read.All', 'TeamMember.Read.All') `
                    -TenantId $TenantId `
                    -UseDeviceCode:$UseDeviceCode `
                    -ForceReconnect:$ForceReconnect `
                    -AutoInstallModules:$AutoInstallModules `
                    -Verbose:$VerbosePreference

                $requiredModules = @(
                'Microsoft.Graph.Groups',
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
            'MicrosoftTeams' {
                Write-Verbose "Connessione a Microsoft Teams..."
                Connect-ToMicrosoftTeams `
                    -TenantId $TenantId `
                    -UseDeviceCode:$UseDeviceCode `
                    -ForceReconnect:$ForceReconnect `
                    -AutoInstallModules:$AutoInstallModules `
                    -Verbose:$VerbosePreference
            }
            default {
                throw "Modulo selezionato non valido: '$SelectedModule'. Moduli validi: $SelectedModule.ValidSet -join ', '."
            }
        }
    }

    process {
        foreach ($currentUser in $User) {
            if ($null -eq $currentUser) {
                continue
            }

            if ($currentUser.PSObject.Properties['Found'] -and -not $currentUser.Found) {
                Write-Warning "Utente non risolto: '$($currentUser.SourceIdentity)'."
                continue
            }

            switch ($SelectedModule) {
                'Microsoft.Graph' {
                        if ([string]::IsNullOrWhiteSpace($currentUser.ObjectId)) {
                        Write-Warning "ObjectId mancante per l'utente '$($currentUser.UserPrincipalName)'."
                        continue
                    }
                }
                'MicrosoftTeams' {
                    if ([string]::IsNullOrWhiteSpace($currentUser.UserPrincipalName)) {
                        Write-Warning "UserPrincipalName mancante per l'utente '$($currentUser.DisplayName)'."
                        continue
                    }
                }
            }
            

            Write-Verbose "Recupero Team per l'utente '$($currentUser.UserPrincipalName)'..."

            try {
                switch($SelectedModule) {
                    'Microsoft.Graph' {
                        $joinedTeams = @(Get-MgUserJoinedTeam -UserId $currentUser.ObjectId -All -ErrorAction Stop)
                    }
                    'MicrosoftTeams' {
                        $joinedTeams = @(Get-Team -User $currentUser.UserPrincipalName -ErrorAction Stop)
                    }
                }
            }
            catch {
                Write-Warning "Impossibile recuperare i Team per '$($currentUser.UserPrincipalName)': $($_.Exception.Message)"
                continue
            }

            foreach ($team in $joinedTeams) {
                $teamRole = 'Member'

                try {
                    switch($SelectedModule) {
                        'Microsoft.Graph' {
                            $teamMembers = @(Get-MgTeamMember -TeamId $team.Id -All -ErrorAction Stop)
                            $teamMembership = $teamMembers | Where-Object { $_.UserId -eq $currentUser.ObjectId } | Select-Object -First 1
                        }
                        'MicrosoftTeams' {
                            $teamMembers = @(Get-TeamUser -GroupId $team.GroupId -ErrorAction Stop)
                            $teamMembership = $teamMembers | Where-Object { $_.User -eq $currentUser.UserPrincipalName } | Select-Object -First 1
                        }
                    }

                    if ($teamMembership -and $teamMembership.Roles -contains 'owner') {
                        $teamRole = 'Owner'
                    }
                }
                catch {
                    Write-Warning "Impossibile determinare il ruolo nel Team '$($team.DisplayName)' per '$($currentUser.UserPrincipalName)': $($_.Exception.Message)"
                }

                switch($SelectedModule) {
                    'Microsoft.Graph' {
                        [pscustomobject][ordered]@{
                            SourceIdentity     = $currentUser.SourceIdentity
                            SourceIdentityType = $currentUser.SourceIdentityType
                            UserObjectId       = $currentUser.ObjectId
                            UserPrincipalName  = $currentUser.UserPrincipalName
                            UserDisplayName    = $currentUser.DisplayName
                            UserMailNickName   = $currentUser.MailNickName
                            TeamId             = $team.Id
                            TeamDisplayName    = $team.DisplayName
                            TeamRole           = $teamRole
                        }
                    }
                    'MicrosoftTeams' {
                        [pscustomobject][ordered]@{
                            SourceIdentity     = $currentUser.SourceIdentity
                            SourceIdentityType = $currentUser.SourceIdentityType
                            UserObjectId       = $currentUser.ObjectId
                            UserPrincipalName  = $currentUser.UserPrincipalName
                            UserDisplayName    = $currentUser.DisplayName
                            UserMailNickName   = $currentUser.MailNickName
                            TeamId             = $team.GroupId
                            TeamDisplayName    = $team.DisplayName
                            TeamRole           = $teamRole
                        }
                    }
                }
            }
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}