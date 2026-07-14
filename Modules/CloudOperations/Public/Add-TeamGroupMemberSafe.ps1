function Add-TeamGroupMemberSafe {
    <#
    .SYNOPSIS
    Aggiunge in modo sicuro un utente a un Microsoft 365 Group associato a un Team,
    gestendo anche il ruolo Member o Owner.

    .DESCRIPTION
    Riceve un Team/Group ID e un utente già risolto, quindi tenta l'aggiunta
    tramite Microsoft Graph.

    Se il ruolo richiesto è Owner:
    - aggiunge l'utente come membro, se necessario;
    - promuove poi l'utente a owner del gruppo sottostante al Team.

    Se il ruolo non è specificato o non riconosciuto, viene usato Member.

    .PARAMETER TeamId
    ID del Team da aggiornare. Coincide con il GroupId del Microsoft 365 Group sottostante.

    .PARAMETER User
    Oggetto utente canonico già risolto.

    .PARAMETER Role
    Ruolo da assegnare. Valori supportati: Member, Owner.
    Default: Member.

    .PARAMETER TenantId
    Tenant ID da usare per la connessione a Microsoft Graph.

    .PARAMETER UseDeviceCode
    Usa l'autenticazione device code.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti.

    .PARAMETER ForceReconnect
    Forza una nuova connessione a Microsoft Graph.

    .PARAMETER Scopes
    Scope Microsoft Graph da usare per la connessione.
    Default: Group.Read.All, GroupMember.ReadWrite.All, User.Read.All.

    .OUTPUTS
    PSCustomObject
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $TeamId,

        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNull()]
        [object[]] $User,

        [Parameter()]
        [ValidateSet('Member', 'Owner')]
        [string] $Role = 'Member',

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $ForceReconnect,

        [Parameter()]
        [string[]] $Scopes = @('Group.Read.All', 'GroupMember.ReadWrite.All', 'User.Read.All')
    )

    begin {
        if (-not [guid]::TryParse($TeamId, [ref]([guid]::Empty))) {
            throw "Il valore specificato per TeamId non è un GUID valido: $TeamId"
        }

        Write-Verbose "Connessione a Microsoft Graph..."
        Connect-ToGraph `
            -Scopes $Scopes `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -ForceReconnect:$ForceReconnect `
            -AutoInstallModules:$AutoInstallModules `
            -Verbose:$VerbosePreference

        $requiredModules = @(
            'Microsoft.Graph.Groups'
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

        try {
            $teamGroup = Get-MgGroup `
                -GroupId $TeamId `
                -Property Id,DisplayName,ResourceProvisioningOptions `
                -ErrorAction Stop

            if ($teamGroup.ResourceProvisioningOptions -notcontains 'Team') {
                throw "L'oggetto '$TeamId' esiste ma non risulta un Team."
            }
        }
        catch {
            throw "Impossibile validare il Team '$TeamId': $($_.Exception.Message)"
        }
    }

    process {
        foreach ($currentUser in $User) {
            if ($null -eq $currentUser) {
                continue
            }

            if ($currentUser.PSObject.Properties['Found'] -and -not $currentUser.Found) {
                [pscustomobject][ordered]@{
                    TeamId            = $teamGroup.Id
                    TeamDisplayName   = $teamGroup.DisplayName
                    RequestedRole     = $Role
                    EffectiveRole     = 'Member'
                    UserObjectId      = $null
                    UserPrincipalName = $null
                    UserDisplayName   = $null
                    UserMailNickName  = $null
                    Status            = 'UserNotFound'
                    Message           = "Utente non risolto: '$($currentUser.SourceIdentity)'."
                }
                continue
            }

            if ([string]::IsNullOrWhiteSpace($currentUser.ObjectId)) {
                [pscustomobject][ordered]@{
                    TeamId            = $teamGroup.Id
                    TeamDisplayName   = $teamGroup.DisplayName
                    RequestedRole     = $Role
                    EffectiveRole     = 'Member'
                    UserObjectId      = $null
                    UserPrincipalName = $currentUser.UserPrincipalName
                    UserDisplayName   = $currentUser.DisplayName
                    UserMailNickName  = $currentUser.MailNickName
                    Status            = 'InvalidUser'
                    Message           = 'ObjectId mancante nell''oggetto utente.'
                }
                continue
            }

            $memberStatus = 'Added'
            $memberMessage = 'Utente aggiunto correttamente al Team.'

            try {
                $directoryObjectRef = @{
                    '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($currentUser.ObjectId)"
                }

                New-MgGroupMemberByRef `
                    -GroupId $teamGroup.Id `
                    -BodyParameter $directoryObjectRef `
                    -ErrorAction Stop
            }
            catch {
                $memberMessage = $_.Exception.Message
                if ($memberMessage -match 'added object references already exist' -or
                    $memberMessage -match 'One or more added object references already exist') {
                    $memberStatus = 'AlreadyMember'
                }
                else {
                    [pscustomobject][ordered]@{
                        TeamId            = $teamGroup.Id
                        TeamDisplayName   = $teamGroup.DisplayName
                        RequestedRole     = $Role
                        EffectiveRole     = 'Member'
                        UserObjectId      = $currentUser.ObjectId
                        UserPrincipalName = $currentUser.UserPrincipalName
                        UserDisplayName   = $currentUser.DisplayName
                        UserMailNickName  = $currentUser.MailNickName
                        Status            = 'AddFailed'
                        Message           = $memberMessage
                    }
                    continue
                }
            }

            if ($Role -eq 'Owner') {
                try {
                    $ownerRef = @{
                        '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($currentUser.ObjectId)"
                    }

                    New-MgGroupOwnerByRef `
                        -GroupId $teamGroup.Id `
                        -BodyParameter $ownerRef `
                        -ErrorAction Stop

                    [pscustomobject][ordered]@{
                        TeamId            = $teamGroup.Id
                        TeamDisplayName   = $teamGroup.DisplayName
                        RequestedRole     = $Role
                        EffectiveRole     = 'Owner'
                        UserObjectId      = $currentUser.ObjectId
                        UserPrincipalName = $currentUser.UserPrincipalName
                        UserDisplayName   = $currentUser.DisplayName
                        UserMailNickName  = $currentUser.MailNickName
                        Status            = if ($memberStatus -eq 'AlreadyMember') { 'OwnerAssigned' } else { 'AddedAsOwner' }
                        Message           = 'Utente aggiunto e impostato come owner del Team.'
                    }
                }
                catch {
                    $ownerMessage = $_.Exception.Message

                    if ($ownerMessage -match 'added object references already exist' -or
                        $ownerMessage -match 'One or more added object references already exist') {
                        [pscustomobject][ordered]@{
                            TeamId            = $teamGroup.Id
                            TeamDisplayName   = $teamGroup.DisplayName
                            RequestedRole     = $Role
                            EffectiveRole     = 'Owner'
                            UserObjectId      = $currentUser.ObjectId
                            UserPrincipalName = $currentUser.UserPrincipalName
                            UserDisplayName   = $currentUser.DisplayName
                            UserMailNickName  = $currentUser.MailNickName
                            Status            = 'AlreadyOwner'
                            Message           = 'Utente già owner del Team.'
                        }
                    }
                    else {
                        [pscustomobject][ordered]@{
                            TeamId            = $teamGroup.Id
                            TeamDisplayName   = $teamGroup.DisplayName
                            RequestedRole     = $Role
                            EffectiveRole     = 'Member'
                            UserObjectId      = $currentUser.ObjectId
                            UserPrincipalName = $currentUser.UserPrincipalName
                            UserDisplayName   = $currentUser.DisplayName
                            UserMailNickName  = $currentUser.MailNickName
                            Status            = 'OwnerAssignFailed'
                            Message           = $ownerMessage
                        }
                    }
                }

                continue
            }

            [pscustomobject][ordered]@{
                TeamId            = $teamGroup.Id
                TeamDisplayName   = $teamGroup.DisplayName
                RequestedRole     = $Role
                EffectiveRole     = 'Member'
                UserObjectId      = $currentUser.ObjectId
                UserPrincipalName = $currentUser.UserPrincipalName
                UserDisplayName   = $currentUser.DisplayName
                UserMailNickName  = $currentUser.MailNickName
                Status            = $memberStatus
                Message           = $memberMessage
            }
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}