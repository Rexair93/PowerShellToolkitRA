function Add-TeamsMembershipFromSpreadsheet {
    <#
    .SYNOPSIS
    Aggiunge utenti a Team o canali privati Microsoft Teams leggendo i dati da un file CSV o XLSX.

    .DESCRIPTION
    Importa un file tabellare tramite Import-SpreadsheetSafe e processa ogni riga come
    istruzione di membership verso un Team oppure un canale privato.

    Supporta anche direttamente il report prodotto da
    Export-UserTeamsAndPrivateChannelsReport.ps1. In questo caso usa
    UserPrincipalName, TeamRole, ChannelRole e ChannelId per ricavare automaticamente
    User, Role e TargetType.

    La funzione riusa le funzioni del toolkit per:
    - selezione file opzionale
    - importazione CSV/XLSX
    - risoluzione identità utente in Entra ID
    - connessione a Microsoft Teams

    Colonne attese nel file:
    - TargetType          : Team oppure PrivateChannel
    - TeamId              : GUID del Team
    - User                : UPN, ObjectId o MailNickName
    - Role                : Member oppure Owner

    Per TargetType = PrivateChannel è richiesta anche:
    - ChannelDisplayName  : nome del canale privato

    .PARAMETER Path
    Percorso del file CSV o XLSX da importare.

    .PARAMETER WorksheetName
    Nome del foglio da leggere, se il file è XLSX.

    .PARAMETER Delimiter
    Delimitatore CSV. Default: ','.

    .PARAMETER PromptForFile
    Se specificato, apre il selettore file quando Path non è valorizzato.

    .PARAMETER UserColumn
    Nome della colonna contenente l'identità utente. Default: User.

    .PARAMETER TargetTypeColumn
    Nome della colonna contenente il tipo destinazione. Default: TargetType.

    .PARAMETER TeamIdColumn
    Nome della colonna contenente il Team ID. Default: TeamId.

    .PARAMETER TeamDisplayNameColumn
    Nome della colonna contenente il nome del Team. Default: TeamDisplayName.

    .PARAMETER ChannelDisplayNameColumn
    Nome della colonna contenente il nome del canale privato. Default: ChannelDisplayName.

    .PARAMETER RoleColumn
    Nome della colonna contenente il ruolo. Default: Role.

    .PARAMETER IdentityType
    Tipo di identità in input: Auto, UserPrincipalName, ObjectId, MailNickName.

    .PARAMETER TenantId
    Tenant ID da usare per la connessione ai servizi Microsoft.

    .PARAMETER UseDeviceCode
    Usa autenticazione device code.

    .PARAMETER AutoInstallModules
    Installa automaticamente i moduli mancanti richiesti.

    .PARAMETER AllowClobber
    Consente AllowClobber durante installazione moduli.

    .PARAMETER ForceReconnect
    Forza una nuova connessione ai servizi richiesti.

    .PARAMETER DelayMilliseconds
    Millisecondi di attesa tra le aggiunte ai Team e quelle ai canali privati. Default: 200.

    .OUTPUTS
    PSCustomObject con:
    RowNumber, TargetType, TeamId, TeamDisplayName, ChannelDisplayName, SourceIdentity, SourceIdentityType,
    ResolvedObjectId, ResolvedUserPrincipalName, Role, Success, Message

    .EXAMPLE
    Add-TeamsMembershipFromSpreadsheet -Path 'C:\Temp\members.xlsx' -Verbose

    .EXAMPLE
    Add-TeamsMembershipFromSpreadsheet -Path 'C:\Temp\members.csv' -Delimiter ';' -WhatIf

    .EXAMPLE
    Add-TeamsMembershipFromSpreadsheet -PromptForFile -UseDeviceCode -Verbose
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
    param(
        [Parameter()]
        [string] $Path,

        [Parameter()]
        [string] $WorksheetName,

        [Parameter()]
        [char] $Delimiter = ',',

        [Parameter()]
        [switch] $PromptForFile,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $UserColumn = 'User',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $TargetTypeColumn = 'TargetType',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $TeamIdColumn = 'TeamId',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $TeamDisplayNameColumn = 'TeamDisplayName',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $ChannelDisplayNameColumn = 'ChannelDisplayName',

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string] $RoleColumn = 'Role',

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
        [switch] $ForceReconnect,

        [Parameter()]
        [ValidateRange(0, [int]::MaxValue)]
        [int] $DelayMilliseconds = 200
    )

    begin {
        $reportInput = $false

        if (-not $Path -and $PromptForFile) {
            $Path = Get-InputFile -Formats @('csv', 'xlsx') -Title 'Seleziona file utenti Teams' -UseConsole:$false
        }

        if ([string]::IsNullOrWhiteSpace($Path)) {
            throw "Specificare -Path oppure usare -PromptForFile."
        }

        Write-Verbose "Connessione a Microsoft Teams..."
        Connect-ToMicrosoftTeams `
            -TenantId            $TenantId `
            -UseDeviceCode:      $UseDeviceCode `
            -ForceReconnect:     $ForceReconnect `
            -AutoInstallModules: $AutoInstallModules `
            -ImportMode          All `
            -Verbose:            $VerbosePreference

        Write-Verbose "Importazione file '$Path'..."
        $rows = @(
            Import-SpreadsheetSafe `
                -Path          $Path `
                -Delimiter     $Delimiter `
                -WorksheetName $WorksheetName
        )

        if (-not $rows -or $rows.Count -eq 0) {
            Write-Warning "Il file '$Path' non contiene righe da elaborare."
            $rows = @()
        }

        if ($rows.Count -gt 0) {
            $availableColumns = @($rows[0].PSObject.Properties.Name)

            if ('User' -notin $availableColumns -and 'UserPrincipalName' -in $availableColumns) {
                $UserColumn = 'UserPrincipalName'
            }

            $reportInput = @('UserPrincipalName', 'TeamRole', 'ChannelRole', 'ChannelId') |
                Where-Object { $_ -in $availableColumns } |
                Measure-Object |
                Select-Object -ExpandProperty Count

            $reportInput = $reportInput -ge 2

            $requiredColumns = @($TeamIdColumn, $UserColumn)
            if (-not $reportInput) {
                $requiredColumns += @($TargetTypeColumn, $RoleColumn)
            }

            foreach ($requiredColumn in $requiredColumns) {
                if ($requiredColumn -notin $availableColumns) {
                    throw "Colonna obbligatoria non trovata: '$requiredColumn'. Colonne disponibili: $($availableColumns -join ', ')"
                }
            }

            $rowsToProcess = @(
                for ($index = 0; $index -lt $rows.Count; $index++) {
                    $row = $rows[$index]
                    $isPrivateChannel = if ($reportInput) {
                        -not [string]::IsNullOrWhiteSpace([string] $row.ChannelId) -or
                        -not [string]::IsNullOrWhiteSpace([string] $row.ChannelRole)
                    }
                    else {
                        ([string] $row.$TargetTypeColumn).Trim().ToLowerInvariant() -eq 'privatechannel'
                    }

                    if ($reportInput -and $isPrivateChannel) {
                        [pscustomobject]@{
                            RowNumber            = $index + 1
                            Row                  = $row
                            IsPrivateChannel     = $false
                            TargetTypeOverride   = 'Team'
                            PrerequisiteRowNumber = $null
                        }
                        [pscustomobject]@{
                            RowNumber            = $index + 1
                            Row                  = $row
                            IsPrivateChannel     = $true
                            TargetTypeOverride   = 'PrivateChannel'
                            PrerequisiteRowNumber = $index + 1
                        }
                    }
                    else {
                        [pscustomobject]@{
                            RowNumber            = $index + 1
                            Row                  = $row
                            IsPrivateChannel     = $isPrivateChannel
                            TargetTypeOverride   = $null
                            PrerequisiteRowNumber = $null
                        }
                    }
                }
            ) | Sort-Object IsPrivateChannel, RowNumber, TargetTypeOverride
        }
        else {
            $rowsToProcess = @()
        }
    }

    process {
        $hasTeamRows = @($rowsToProcess | Where-Object { -not $_.IsPrivateChannel }).Count -gt 0
        $hasPrivateChannelRows = @($rowsToProcess | Where-Object IsPrivateChannel).Count -gt 0
        $delayApplied = $false
        $teamResults = @{}

        foreach ($rowToProcess in $rowsToProcess) {
            if ($hasTeamRows -and $hasPrivateChannelRows -and
                $rowToProcess.IsPrivateChannel -and -not $delayApplied) {
                Write-Verbose "Attesa di $DelayMilliseconds ms prima di elaborare i canali privati..."
                if ($DelayMilliseconds -gt 0) {
                    Start-Sleep -Milliseconds $DelayMilliseconds
                }
                $delayApplied = $true
            }

            $rowNumber = $rowToProcess.RowNumber
            $row = $rowToProcess.Row

            $targetType         = if ($null -ne $rowToProcess.TargetTypeOverride) {
                $rowToProcess.TargetTypeOverride
            }
            elseif ($reportInput) {
                if (-not [string]::IsNullOrWhiteSpace([string] $row.ChannelId) -or
                    -not [string]::IsNullOrWhiteSpace([string] $row.ChannelRole)) {
                    'PrivateChannel'
                }
                else {
                    'Team'
                }
            }
            else {
                [string] $row.$TargetTypeColumn
            }
            $teamId             = [string] $row.$TeamIdColumn
            $teamDisplayName    = if ($reportInput) { [string] $row.TeamDisplayName } else { [string] $row.$TeamDisplayNameColumn }
            $channelDisplayName = if ($targetType -eq 'PrivateChannel') { [string] $row.$ChannelDisplayNameColumn } else { $null }
            $sourceIdentity     = if ($reportInput) { [string] $row.UserPrincipalName } else { [string] $row.$UserColumn }
            $role               = if ($reportInput) {
                if ($targetType -eq 'PrivateChannel') { [string] $row.ChannelRole } else { [string] $row.TeamRole }
            }
            else {
                [string] $row.$RoleColumn
            }

            $result = [pscustomobject][ordered]@{
                RowNumber                 = $rowNumber
                TargetType                = $targetType
                TeamId                    = $teamId
                TeamDisplayName           = $teamDisplayName
                ChannelDisplayName        = $channelDisplayName
                SourceIdentity            = $sourceIdentity
                SourceIdentityType        = $null
                ResolvedObjectId          = $null
                ResolvedUserPrincipalName = $null
                Role                      = $role
                Success                   = $false
                Message                   = $null
            }

            try {
                if ($null -ne $rowToProcess.PrerequisiteRowNumber -and
                    $false -eq $teamResults[[string] $rowToProcess.PrerequisiteRowNumber]) {
                    throw "Prerequisito Team della riga $($rowToProcess.PrerequisiteRowNumber) non completato; canale privato non elaborato."
                }

                if ([string]::IsNullOrWhiteSpace($targetType)) {
                    throw "Colonna '$TargetTypeColumn' mancante o vuota."
                }

                if ([string]::IsNullOrWhiteSpace($teamId)) {
                    throw "Colonna '$TeamIdColumn' mancante o vuota."
                }

                if (-not [guid]::TryParse($teamId, [ref]([guid]::Empty))) {
                    throw "Il TeamId '$teamId' non è un GUID valido."
                }

                if ([string]::IsNullOrWhiteSpace($sourceIdentity)) {
                    throw "Colonna '$UserColumn' mancante o vuota."
                }

                if ([string]::IsNullOrWhiteSpace($role)) {
                    throw "Colonna '$RoleColumn' mancante o vuota."
                }

                switch ($targetType.Trim().ToLowerInvariant()) {
                    'team' {
                        $normalizedTargetType = 'Team'
                    }
                    'privatechannel' {
                        $normalizedTargetType = 'PrivateChannel'
                    }
                    default {
                        throw "Valore TargetType non supportato: '$targetType'. Valori ammessi: Team, PrivateChannel."
                    }
                }

                switch ($role.Trim().ToLowerInvariant()) {
                    'member' {
                        $normalizedRole = 'Member'
                    }
                    'owner' {
                        $normalizedRole = 'Owner'
                    }
                    default {
                        throw "Valore Role non supportato: '$role'. Valori ammessi: Member, Owner."
                    }
                }

                if ($normalizedTargetType -eq 'PrivateChannel' -and [string]::IsNullOrWhiteSpace($channelDisplayName)) {
                    throw "Per TargetType 'PrivateChannel' la colonna '$ChannelDisplayNameColumn' è obbligatoria."
                }

                $resolvedUser = @(
                    Resolve-EntraUserIdentity `
                        -Identity            $sourceIdentity `
                        -IdentityType        $IdentityType `
                        -TenantId            $TenantId `
                        -UseDeviceCode:      $UseDeviceCode `
                        -AutoInstallModules: $AutoInstallModules `
                        -AllowClobber:       $AllowClobber `
                        -ForceReconnect:     $ForceReconnect `
                        -Verbose:            $VerbosePreference
                ) | Select-Object -First 1

                if ($null -eq $resolvedUser -or -not $resolvedUser.Found) {
                    throw "Impossibile risolvere l'identità utente '$sourceIdentity'."
                }

                if ([string]::IsNullOrWhiteSpace($resolvedUser.UserPrincipalName)) {
                    throw "L'identità '$sourceIdentity' è stata risolta ma non contiene un UserPrincipalName valido."
                }

                $result.SourceIdentityType        = $resolvedUser.SourceIdentityType
                $result.ResolvedObjectId          = $resolvedUser.ObjectId
                $result.ResolvedUserPrincipalName = $resolvedUser.UserPrincipalName
                $result.TargetType                = $normalizedTargetType
                $result.Role                      = $normalizedRole

                if ($normalizedTargetType -eq 'Team') {
                    $existingMembership = @(Get-TeamUser -GroupId $teamId -ErrorAction Stop) |
                        Where-Object {
                            ($_.UserId -and $_.UserId -eq $resolvedUser.ObjectId) -or
                            ($_.User -and $_.User -ieq $resolvedUser.UserPrincipalName)
                        } |
                        Select-Object -First 1
                    $existingRole = if ($existingMembership -and ([string] $existingMembership.Role).Trim() -ieq 'Owner') { 'Owner' } else { 'Member' }
                    $targetName = if ([string]::IsNullOrWhiteSpace($teamDisplayName)) { $teamId } else { $teamDisplayName }

                    if ($existingMembership -and $existingRole -ieq $normalizedRole) {
                        $result.Success = $true
                        $result.Message = "Utente già presente nel Team '$targetName' con ruolo '$existingRole'."
                    }
                    else {
                        $actionDescription = "Aggiunta o aggiornamento dell'utente '$($resolvedUser.UserPrincipalName)' al Team '$teamId' come '$normalizedRole'"

                        if ($PSCmdlet.ShouldProcess($teamId, $actionDescription)) {
                            Add-TeamUser `
                                -GroupId     $teamId `
                                -User        $resolvedUser.UserPrincipalName `
                                -Role        $normalizedRole `
                                -ErrorAction Stop

                            $result.Success = $true
                            if ($existingMembership) {
                                $result.Message = "Ruolo aggiornato nel Team '$targetName' da '$existingRole' a '$normalizedRole'."
                            }
                            else {
                                $result.Message = "Utente aggiunto al Team '$targetName' con ruolo '$normalizedRole'."
                            }
                        }
                        else {
                            $result.Success = $true
                            $result.Message = "Operazione simulata: aggiunta o aggiornamento dell'utente nel Team '$targetName' con ruolo '$normalizedRole'."
                        }
                    }
                }

                if ($normalizedTargetType -eq 'PrivateChannel') {
                    $existingMembership = @(Get-TeamChannelUser `
                            -GroupId     $teamId `
                            -DisplayName $channelDisplayName `
                            -ErrorAction Stop) |
                        Where-Object {
                            ($_.UserId -and $_.UserId -eq $resolvedUser.ObjectId) -or
                            ($_.User -and $_.User -ieq $resolvedUser.UserPrincipalName)
                        } |
                        Select-Object -First 1
                    $existingRole = if ($existingMembership -and ([string] $existingMembership.Role).Trim() -ieq 'Owner') { 'Owner' } else { 'Member' }

                    if ($existingMembership -and $existingRole -ieq $normalizedRole) {
                        $result.Success = $true
                        $result.Message = "Utente già presente nel canale privato '$channelDisplayName' con ruolo '$existingRole'."
                    }
                    else {
                        $wasAdded = $false
                        $wasPromoted = $false

                        if (-not $existingMembership) {
                            $actionDescription = "Aggiunta dell'utente '$($resolvedUser.UserPrincipalName)' al canale privato '$channelDisplayName' del Team '$teamId'"

                            if ($PSCmdlet.ShouldProcess($channelDisplayName, $actionDescription)) {
                                Add-TeamChannelUser `
                                    -GroupId     $teamId `
                                    -DisplayName $channelDisplayName `
                                    -User        $resolvedUser.UserPrincipalName `
                                    -ErrorAction Stop

                                $wasAdded = $true
                            }
                        }

                        if ($normalizedRole -eq 'Owner' -and
                            ($null -ne $existingMembership -or $wasAdded) -and
                            (-not $existingMembership -or $existingRole -ine 'Owner')) {
                            $actionDescription = "Promozione dell'utente '$($resolvedUser.UserPrincipalName)' a Owner del canale privato '$channelDisplayName'"

                            if ($PSCmdlet.ShouldProcess($channelDisplayName, $actionDescription)) {
                                Add-TeamChannelUser `
                                    -GroupId     $teamId `
                                    -DisplayName $channelDisplayName `
                                    -User        $resolvedUser.UserPrincipalName `
                                    -Role        Owner `
                                    -ErrorAction Stop

                                $wasPromoted = $true
                            }
                        }

                        $result.Success = $true
                        if ($wasPromoted -and $wasAdded) {
                            $result.Message = "Utente aggiunto al canale privato '$channelDisplayName' e promosso a Owner."
                        }
                        elseif ($wasPromoted) {
                            $result.Message = "Ruolo aggiornato nel canale privato '$channelDisplayName' da '$existingRole' a 'Owner'."
                        }
                        elseif ($wasAdded) {
                            $result.Message = "Utente aggiunto al canale privato '$channelDisplayName' con ruolo 'Member'."
                        }
                        else {
                            $result.Message = "Operazione simulata: aggiunta dell'utente al canale privato '$channelDisplayName' e, se richiesto, promozione a Owner."
                        }
                    }
                }

            }
            catch {
                $result.Message = $_.Exception.Message
                Write-Warning "Riga ${rowNumber}: $($result.Message)"
            }

            if ($targetType -eq 'Team') {
                $teamResults[[string] $rowNumber] = $result.Success
            }

            $result
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}