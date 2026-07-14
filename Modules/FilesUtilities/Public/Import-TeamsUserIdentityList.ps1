function Import-TeamsUserIdentityList {
    <#
    .SYNOPSIS
    Importa una lista utenti da CSV o XLSX e restituisce identità normalizzate.

    .DESCRIPTION
    Legge un file tabellare tramite Import-SpreadsheetSafe e individua la colonna
    contenente l'identificativo utente. Supporta UPN, ObjectId o MailNickName.

    Se presente una colonna ruolo denominata 'Ruolo' o 'Role', il valore viene
    normalizzato ai ruoli supportati:
    - Owner
    - Member

    Se la colonna non è presente, è vuota o contiene un valore non riconosciuto,
    il ruolo restituito sarà 'Member'.

    Restituisce oggetti con:
    RowNumber, SourceIdentity, SourceIdentityType, RequestedRole.

    .PARAMETER Path
    Percorso del file di input.

    .PARAMETER WorksheetName
    Nome del foglio da leggere, se il file è XLSX.

    .PARAMETER IdentityColumn
    Nome esplicito della colonna da usare. Se omesso, viene tentato il riconoscimento automatico.

    .OUTPUTS
    PSCustomObject
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,

        [Parameter()]
        [string] $WorksheetName,

        [Parameter()]
        [string] $IdentityColumn,

        [Parameter()]
        [char] $Delimiter
    )

    function Get-AutoDetectedDelimiter {
        param(
            [Parameter(Mandatory)]
            [string] $CsvPath
        )

        $firstLine = Get-Content -Path $CsvPath -TotalCount 1
        if ([string]::IsNullOrWhiteSpace($firstLine)) {
            return ','
        }

        $semicolonCount = ([regex]::Matches($firstLine, ';')).Count
        $commaCount     = ([regex]::Matches($firstLine, ',')).Count
        $tabCount       = ([regex]::Matches($firstLine, "`t")).Count

        if ($semicolonCount -ge $commaCount -and $semicolonCount -ge $tabCount -and $semicolonCount -gt 0) {
            return ';'
        }

        if ($tabCount -ge $commaCount -and $tabCount -gt 0) {
            return "`t"
        }

        return ','
    }

    $effectiveDelimiter = $Delimiter
    $extension = ConvertTo-NormalizedExt ([IO.Path]::GetExtension($Path))

    if (-not $PSBoundParameters.ContainsKey('Delimiter') -and $extension -eq 'csv') {
        $effectiveDelimiter = Get-AutoDetectedDelimiter -CsvPath $Path
    }

    $importParams = @{
        Path = $Path
    }

    if ($WorksheetName) {
        $importParams['WorksheetName'] = $WorksheetName
    }

    if ($extension -eq 'csv' -and $effectiveDelimiter) {
        $importParams['Delimiter'] = $effectiveDelimiter
    }

    $rows = Import-SpreadsheetSafe @importParams
    if (-not $rows) {
        return @()
    }

    $candidateColumns = @(
        'UserPrincipalName', 'UPN', 'ObjectId', 'Id', 'MailNickName', 'MailNickname', 'Identity', 'User'
    )

    $roleCandidateColumns = @(
        'Ruolo', 'Role'
    )

    if (-not $IdentityColumn) {
        $availableColumns = @($rows[0].PSObject.Properties.Name)

        $IdentityColumn = $candidateColumns |
            Where-Object { $_ -in $availableColumns } |
            Select-Object -First 1

        if (-not $IdentityColumn) {
            throw "Nessuna colonna identità riconosciuta. Colonne disponibili: $($availableColumns -join ', ')"
        }
    }

    $availableColumns = @($rows[0].PSObject.Properties.Name)

    $roleColumn = $roleCandidateColumns |
        Where-Object { $_ -in $availableColumns } |
        Select-Object -First 1

    function Resolve-RequestedRole {
        param(
            [AllowNull()]
            [string] $Value
        )

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return 'Member'
        }

        switch ($Value.Trim().ToLowerInvariant()) {
            'owner'         { 'Owner'; break }
            'owners'        { 'Owner'; break }
            'proprietario'  { 'Owner'; break }
            'proprietari'   { 'Owner'; break }
            'member'        { 'Member'; break }
            'members'       { 'Member'; break }
            'membro'        { 'Member'; break }
            'membri'        { 'Member'; break }
            default         { 'Member' }
        }
    }

    $rowNumber = 0

    foreach ($row in $rows) {
        $rowNumber++

        $value = [string] $row.$IdentityColumn
        if ([string]::IsNullOrWhiteSpace($value)) {
            continue
        }

        $value = $value.Trim()

        $identityType = if ($IdentityColumn -match '^(UserPrincipalName|UPN)$' -or $value -match '@') {
            'UserPrincipalName'
        }
        elseif ($IdentityColumn -match '^(ObjectId|Id)$' -or [guid]::TryParse($value, [ref]([guid]::Empty))) {
            'ObjectId'
        }
        else {
            'MailNickName'
        }

        $requestedRole = if ($roleColumn) {
            Resolve-RequestedRole -Value ([string] $row.$roleColumn)
        }
        else {
            'Member'
        }

        [pscustomobject][ordered]@{
            RowNumber          = $rowNumber
            SourceIdentity     = $value
            SourceIdentityType = $identityType
            RequestedRole      = $requestedRole
        }
    }
}