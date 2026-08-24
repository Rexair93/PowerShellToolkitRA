[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter()]
    [string] $InputPath,

    [Parameter()]
    [string] $WorksheetName,

    [Parameter()]
    [string] $IdentityColumn,

    [Parameter()]
    [string] $TenantId,

    [Parameter()]
    [ValidateSet('csv', 'xlsx')]
    [string] $OutputFormat = 'xlsx',

    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [string] $SelectedModule,

    [Parameter()]
    [switch] $UseConsole,

    [Parameter()]
    [switch] $UseDeviceCode,

    [Parameter()]
    [switch] $AutoInstallModules,

    [Parameter()]
    [switch] $AllowClobber,

    [Parameter()]
    [switch] $ForceReconnect,

    [Parameter()]
    [switch] $IncludeTeamsWithoutPrivateChannels
)

begin {
    $ErrorActionPreference = 'Stop'
    
    Import-Module CloudOperations -ErrorAction Stop
    Import-Module FilesUtilities  -ErrorAction Stop

    function Export-TeamsMembershipResults {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [ValidateNotNull()]
            [object[]] $InputObject,

            [Parameter(Mandatory)]
            [ValidateNotNullOrEmpty()]
            [string] $Path,

            [Parameter(Mandatory)]
            [ValidateSet('csv', 'xlsx')]
            [string] $Format
        )

        switch ($Format) {
            'csv' {
                $InputObject |
                    Export-Csv `
                        -Path $Path `
                        -NoTypeInformation `
                        -Encoding UTF8
            }

            'xlsx' {
                if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
                    throw "Per esportare in XLSX è richiesto il modulo 'ImportExcel'."
                }

                if (-not (Get-Module -Name ImportExcel)) {
                    Import-Module ImportExcel -ErrorAction Stop -Verbose:$false | Out-Null
                }

                $InputObject |
                    Export-Excel `
                        -Path $Path `
                        -WorksheetName 'Memberships' `
                        -TableName 'Memberships' `
                        -AutoSize `
                        -FreezeTopRow `
                        -BoldTopRow `
                        -ClearSheet
            }
        }
    }

    Write-Verbose "Avvio script di orchestrazione report Teams..."
}

process {
    if (-not $SelectedModule) {
        Write-Host ""
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host " Export User Teams and Private Channels Report" -ForegroundColor Cyan
        Write-Host "============================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Seleziona modulo da utilizzare:"
        Write-Host "1 - Only Microsoft Graph Module"
        Write-Host "2 - Both Microsoft Graph and Microsoft Teams Module"
        Write-Host ""

        $moduleChoice = Read-Host "Inserisci la scelta (1 o 2)"

        switch ($moduleChoice) {
            '1' { $SelectedModule = 'Microsoft.Graph' }
            '2' { $SelectedModule = 'MicrosoftTeams' }
            default { throw "Scelta non valida. Inserire 1 oppure 2." }
        }
    }

    if (-not $InputPath) {
        $InputPath = Get-InputFile `
            -Formats @('csv', 'xlsx') `
            -Title 'Seleziona il file con la lista utenti' `
            -UseConsole:$UseConsole
    }

    if (-not $OutputPath) {
        $destination = Get-ExportDestination `
            -DefaultFileName 'teams-membership-report.xlsx' `
            -Formats @('xlsx', 'csv') `
            -PreferredFormat $OutputFormat `
            -Title 'Scegli dove salvare il report' `
            -UseConsole:$UseConsole

        $OutputPath = $destination.Path
        $OutputFormat = $destination.Format
    }

    Write-Verbose "Import lista utenti da '$InputPath'..."
    $inputUsers = @(Import-TeamsUserIdentityList `
        -Path $InputPath `
        -WorksheetName $WorksheetName `
        -IdentityColumn $IdentityColumn `
        -Verbose:$VerbosePreference)

    if (-not $inputUsers) {
        throw "Nessun utente valido trovato nel file di input."
    }

    Write-Verbose "Utenti importati: $($inputUsers.Count)."

    $resolvedUsers = foreach ($inputUser in $inputUsers) {
        Resolve-EntraUserIdentity `
            -Identity $inputUser.SourceIdentity `
            -IdentityType $inputUser.SourceIdentityType `
            -TenantId $TenantId `
            -UseDeviceCode:$UseDeviceCode `
            -AutoInstallModules:$AutoInstallModules `
            -AllowClobber:$AllowClobber `
            -ForceReconnect:$ForceReconnect `
            -Verbose:$VerbosePreference
    }

    $resolvedUsers = @($resolvedUsers)

    if (-not $resolvedUsers) {
        throw "Nessun utente risolto."
    }

    $notFoundUsers = @($resolvedUsers | Where-Object { -not $_.Found })
    if ($notFoundUsers.Count -gt 0) {
        Write-Warning "Utenti non risolti: $($notFoundUsers.Count)."
    }

    $validUsers = @($resolvedUsers | Where-Object { $_.Found -and -not [string]::IsNullOrWhiteSpace($_.ObjectId) })
    if (-not $validUsers) {
        throw "Nessun utente validamente risolto con ObjectId disponibile."
    }

    Write-Verbose "Utenti risolti correttamente: $($validUsers.Count)."

    $teamMemberships = @(
        $validUsers |
            Get-UserTeamsMembershipDetail `
                -TenantId $TenantId `
                -SelectedModule $SelectedModule `
                -UseDeviceCode:$UseDeviceCode `
                -AutoInstallModules:$AutoInstallModules `
                -ForceReconnect:$ForceReconnect `
                -Verbose:$VerbosePreference
    )
    

    if (-not $teamMemberships) {
        Write-Warning "Nessun Team trovato per gli utenti specificati."

        if ($IncludeTeamsWithoutPrivateChannels) {
            $results = @()
        }
        else {
            $results = @()
        }
    }
    else {
        Write-Verbose "Relazioni utente-Team trovate: $($teamMemberships.Count)."

        $privateChannelParams = @{
            TenantId            = $TenantId
            SelectedModule      = $SelectedModule
            UseDeviceCode       = $UseDeviceCode
            AutoInstallModules  = $AutoInstallModules
            ForceReconnect      = $ForceReconnect
            Verbose             = ($VerbosePreference -ne 'SilentlyContinue')
        }

        if ($IncludeTeamsWithoutPrivateChannels) {
            $privateChannelParams['IncludeTeamsWithoutPrivateChannels'] = $true
        }

        $results = @(
            $teamMemberships |
                Get-UserPrivateChannelMembershipDetail @privateChannelParams
        )
    }

    $results = @($results)

    if (-not $results) {
        Write-Warning "Nessun risultato finale da esportare."

        $results = @(
            foreach ($user in $validUsers) {
                [pscustomobject][ordered]@{
                    SourceIdentity         = $user.SourceIdentity
                    SourceIdentityType     = $user.SourceIdentityType
                    UserObjectId           = $user.ObjectId
                    UserPrincipalName      = $user.UserPrincipalName
                    UserDisplayName        = $user.DisplayName
                    UserMailNickName       = $user.MailNickName
                    TeamId                 = $null
                    TeamDisplayName        = $null
                    TeamRole               = $null
                    ChannelId              = $null
                    ChannelDisplayName     = $null
                    ChannelMembershipType  = $null
                    ChannelRole            = $null
                    ResolutionFound        = $user.Found
                }
            }
        )
    }
    else {
        $results = foreach ($row in $results) {
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
                ChannelId              = $row.ChannelId
                ChannelDisplayName     = $row.ChannelDisplayName
                ChannelMembershipType  = $row.ChannelMembershipType
                ChannelRole            = $row.ChannelRole
                ResolutionFound        = $true
            }
        }
    }

    if ($PSCmdlet.ShouldProcess($OutputPath, "Esportazione report Teams membership")) {
        Export-TeamsMembershipResults `
            -InputObject $results `
            -Path $OutputPath `
            -Format $OutputFormat
    }

    Write-Information "Report esportato in: $OutputPath"
}

end {
    Write-Verbose "Script completato."
}