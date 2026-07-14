function Add-TeamGroupMemberSafe {
    <#
    .SYNOPSIS
    Aggiunge in modo sicuro un utente a un Microsoft 365 Group associato a un Team.

    .DESCRIPTION
    Riceve un Team/Group ID e un utente già risolto, quindi tenta l'aggiunta
    tramite Microsoft Graph.

    Se il TeamId non corrisponde a un Team valido, genera un errore.
    Se l'utente è già membro, restituisce uno stato specifico senza interrompere il flusso.

    .PARAMETER TeamId
    ID del Team da aggiornare. Coincide con il GroupId del Microsoft 365 Group sottostante.

    .PARAMETER User
    Oggetto utente canonico già risolto, contenente almeno:
    ObjectId, UserPrincipalName, DisplayName, MailNickName, SourceIdentity, SourceIdentityType.

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
    PSCustomObject con:
    TeamId, TeamDisplayName, UserObjectId, UserPrincipalName, UserDisplayName,
    UserMailNickName, Status, Message.

    .EXAMPLE
    $user | Add-TeamGroupMemberSafe -TeamId '00000000-0000-0000-0000-000000000000'

    .EXAMPLE
    Add-TeamGroupMemberSafe `
        -TeamId '00000000-0000-0000-0000-000000000000' `
        -User $resolvedUser `
        -TenantId '11111111-1111-1111-1111-111111111111' `
        -UseDeviceCode
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

        Write-Verbose "Validazione Team target..."
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
                    UserObjectId      = $null
                    UserPrincipalName = $currentUser.UserPrincipalName
                    UserDisplayName   = $currentUser.DisplayName
                    UserMailNickName  = $currentUser.MailNickName
                    Status            = 'InvalidUser'
                    Message           = 'ObjectId mancante nell''oggetto utente.'
                }
                continue
            }

            try {
                $directoryObjectRef = @{
                    '@odata.id' = "https://graph.microsoft.com/v1.0/directoryObjects/$($currentUser.ObjectId)"
                }

                New-MgGroupMemberByRef `
                    -GroupId $teamGroup.Id `
                    -BodyParameter $directoryObjectRef `
                    -ErrorAction Stop

                [pscustomobject][ordered]@{
                    TeamId            = $teamGroup.Id
                    TeamDisplayName   = $teamGroup.DisplayName
                    UserObjectId      = $currentUser.ObjectId
                    UserPrincipalName = $currentUser.UserPrincipalName
                    UserDisplayName   = $currentUser.DisplayName
                    UserMailNickName  = $currentUser.MailNickName
                    Status            = 'Added'
                    Message           = 'Utente aggiunto correttamente al Team.'
                }
            }
            catch {
                $message = $_.Exception.Message
                $status = 'AddFailed'

                if ($message -match 'added object references already exist' -or
                    $message -match 'One or more added object references already exist') {
                    $status = 'AlreadyMember'
                }

                [pscustomobject][ordered]@{
                    TeamId            = $teamGroup.Id
                    TeamDisplayName   = $teamGroup.DisplayName
                    UserObjectId      = $currentUser.ObjectId
                    UserPrincipalName = $currentUser.UserPrincipalName
                    UserDisplayName   = $currentUser.DisplayName
                    UserMailNickName  = $currentUser.MailNickName
                    Status            = $status
                    Message           = $message
                }
            }
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}