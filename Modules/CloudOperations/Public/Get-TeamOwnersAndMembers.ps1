function Get-TeamOwnersAndMembers {
    <#
    .SYNOPSIS
    Recupera owners e membri di uno o più Team Microsoft 365 a partire dal solo Team ID.

    .DESCRIPTION
    Accetta una lista di Team ID, recupera via Microsoft Graph il DisplayName del team
    e i relativi owners e membri del gruppo sottostante, quindi restituisce un oggetto
    con le proprietà:
    teamID, teamDisplayName, teamOwners, teamOwnersCount, teamMembers, teamMembersCount.

    L'identità da esportare per owners e members è configurabile tramite IdentityProperty:
    UPN, MailNickName o DisplayName. Il separatore tra i valori è configurabile tramite Separator.

    .PARAMETER TeamId
    Uno o più Team ID (GroupId) da elaborare.

    .PARAMETER IdentityProperty
    Proprietà da usare per rappresentare owners e members nell'output.
    Valori ammessi: UPN, MailNickName, DisplayName.
    Default: UPN.

    .PARAMETER Separator
    Separatore da usare per concatenare owners e members.
    Default: ';'.

    .PARAMETER Scopes
    Scope Microsoft Graph da usare per la connessione.
    Default: Group.Read.All, User.Read.All.

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
    teamID, teamDisplayName, teamOwners, teamOwnersCount, teamMembers, teamMembersCount.

    .EXAMPLE
    Get-TeamOwnersAndMembers -TeamId '00000000-0000-0000-0000-000000000000'

    .EXAMPLE
    '00000000-0000-0000-0000-000000000000','11111111-1111-1111-1111-111111111111' |
        Get-TeamOwnersAndMembers -IdentityProperty DisplayName -Separator '|'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [Alias('GroupId', 'Id')]
        [string[]] $TeamId,

        [Parameter()]
        [ValidateSet('UPN', 'MailNickName', 'DisplayName')]
        [string] $IdentityProperty = 'UPN',

        [Parameter()]
        [string] $Separator = ';',

        [Parameter()]
        [string[]] $Scopes = @('Group.Read.All', 'User.Read.All'),

        [Parameter()]
        [string] $TenantId,

        [Parameter()]
        [switch] $UseDeviceCode,

        [Parameter()]
        [switch] $AutoInstallModules,

        [Parameter()]
        [switch] $ForceReconnect
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

        function Get-DirectoryObjectValue {
            param(
                [Parameter(Mandatory)]
                [object] $Object,

                [Parameter(Mandatory)]
                [string] $PropertyName
            )

            switch ($PropertyName) {
                'UPN' {
                    if ($Object.PSObject.Properties['UserPrincipalName'] -and -not [string]::IsNullOrWhiteSpace($Object.UserPrincipalName)) {
                        return $Object.UserPrincipalName
                    }
                    if ($Object.PSObject.Properties['AdditionalProperties'] -and $Object.AdditionalProperties.ContainsKey('userPrincipalName')) {
                        return [string] $Object.AdditionalProperties['userPrincipalName']
                    }
                }

                'MailNickName' {
                    if ($Object.PSObject.Properties['MailNickname'] -and -not [string]::IsNullOrWhiteSpace($Object.MailNickname)) {
                        return $Object.MailNickname
                    }
                    if ($Object.PSObject.Properties['AdditionalProperties'] -and $Object.AdditionalProperties.ContainsKey('mailNickname')) {
                        return [string] $Object.AdditionalProperties['mailNickname']
                    }
                }

                'DisplayName' {
                    if ($Object.PSObject.Properties['DisplayName'] -and -not [string]::IsNullOrWhiteSpace($Object.DisplayName)) {
                        return $Object.DisplayName
                    }
                    if ($Object.PSObject.Properties['AdditionalProperties'] -and $Object.AdditionalProperties.ContainsKey('displayName')) {
                        return [string] $Object.AdditionalProperties['displayName']
                    }
                }
            }

            return $null
        }
    }

    process {
        foreach ($currentTeamId in $TeamId) {
            if ([string]::IsNullOrWhiteSpace($currentTeamId)) {
                continue
            }

            if (-not [guid]::TryParse($currentTeamId, [ref]([guid]::Empty))) {
                Write-Warning "Il valore '$currentTeamId' non è un GUID valido."
                continue
            }

            Write-Verbose "Elaborazione Team ID '$currentTeamId'..."

            try {
                $group = Get-MgGroup `
                    -GroupId $currentTeamId `
                    -Property Id,DisplayName,ResourceProvisioningOptions `
                    -ErrorAction Stop

                if ($group.ResourceProvisioningOptions -notcontains 'Team') {
                    Write-Warning "L'oggetto '$currentTeamId' esiste ma non risulta un Team."
                    continue
                }

                $owners = @(Get-MgGroupOwner -GroupId $group.Id -All -ErrorAction Stop)
                $members = @(Get-MgGroupMember -GroupId $group.Id -All -ErrorAction Stop)

                $ownerValues = foreach ($owner in $owners) {
                    $value = Get-DirectoryObjectValue -Object $owner -PropertyName $IdentityProperty
                    if (-not [string]::IsNullOrWhiteSpace($value)) {
                        $value
                    }
                }

                $memberValues = foreach ($member in $members) {
                    $value = Get-DirectoryObjectValue -Object $member -PropertyName $IdentityProperty
                    if (-not [string]::IsNullOrWhiteSpace($value)) {
                        $value
                    }
                }

                $ownerValues = @($ownerValues | Sort-Object -Unique)
                $memberValues = @($memberValues | Sort-Object -Unique)

                [pscustomobject][ordered]@{
                    teamID           = $group.Id
                    teamDisplayName  = $group.DisplayName
                    teamOwners       = ($ownerValues -join $Separator)
                    teamOwnersCount  = $ownerValues.Count
                    teamMembers      = ($memberValues -join $Separator)
                    teamMembersCount = $memberValues.Count
                }
            }
            catch {
                Write-Warning "Errore durante l'elaborazione del Team ID '$currentTeamId': $($_.Exception.Message)"
            }
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}