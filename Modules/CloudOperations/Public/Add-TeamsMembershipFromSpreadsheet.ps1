function Add-TeamsMembershipFromSpreadsheet {
    <#
    .SYNOPSIS
    Aggiunge utenti a Team o canali privati Microsoft Teams leggendo i dati da un file CSV o XLSX.

    .DESCRIPTION
    Importa un file tabellare tramite Import-SpreadsheetSafe e processa ogni riga come
    istruzione di membership verso un Team oppure un canale privato.

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

    .OUTPUTS
    PSCustomObject con:
    RowNumber, TargetType, TeamId, ChannelDisplayName, SourceIdentity, SourceIdentityType,
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
        [switch] $ForceReconnect
    )

    begin {
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

            foreach ($requiredColumn in @($TargetTypeColumn, $TeamIdColumn, $UserColumn, $RoleColumn)) {
                if ($requiredColumn -notin $availableColumns) {
                    throw "Colonna obbligatoria non trovata: '$requiredColumn'. Colonne disponibili: $($availableColumns -join ', ')"
                }
            }
        }
    }

    process {
        $rowNumber = 0

        foreach ($row in $rows) {
            $rowNumber++

            $targetType         = [string] $row.$TargetTypeColumn
            $teamId             = [string] $row.$TeamIdColumn
            $channelDisplayName = [string] $row.$ChannelDisplayNameColumn
            $sourceIdentity     = [string] $row.$UserColumn
            $role               = [string] $row.$RoleColumn

            $result = [pscustomobject][ordered]@{
                RowNumber                 = $rowNumber
                TargetType                = $targetType
                TeamId                    = $teamId
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
                    $actionDescription = "Aggiunta utente '$($resolvedUser.UserPrincipalName)' al Team '$teamId' come '$normalizedRole'"

                    if ($PSCmdlet.ShouldProcess($teamId, $actionDescription)) {
                        Add-TeamUser `
                            -GroupId     $teamId `
                            -User        $resolvedUser.UserPrincipalName `
                            -Role        $normalizedRole `
                            -ErrorAction Stop
                    }
                }
                else {
                    $actionDescription = "Aggiunta utente '$($resolvedUser.UserPrincipalName)' al canale privato '$channelDisplayName' del Team '$teamId' come '$normalizedRole'"

                    if ($PSCmdlet.ShouldProcess($channelDisplayName, $actionDescription)) {
                        Add-TeamChannelUser `
                            -GroupId     $teamId `
                            -DisplayName $channelDisplayName `
                            -User        $resolvedUser.UserPrincipalName `
                            -Role        $normalizedRole `
                            -ErrorAction Stop
                    }
                }

                $result.Success = $true
                $result.Message = 'Operazione completata.'
            }
            catch {
                $result.Message = $_.Exception.Message
                Write-Warning "Riga ${rowNumber}: $($result.Message)"
            }

            $result
        }
    }

    end {
        Write-Verbose "Elaborazione completata."
    }
}