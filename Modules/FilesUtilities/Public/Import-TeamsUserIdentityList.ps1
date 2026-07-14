function Import-TeamsUserIdentityList {
    <#
    .SYNOPSIS
    Importa una lista utenti da CSV o XLSX e restituisce identità normalizzate.

    .DESCRIPTION
    Legge un file tabellare tramite Import-SpreadsheetSafe e individua la colonna
    contenente l'identificativo utente. Supporta UPN, ObjectId o MailNickName.

    Restituisce oggetti con:
    RowNumber, SourceIdentity, SourceIdentityType.

    .PARAMETER Path
    Percorso del file di input.

    .PARAMETER WorksheetName
    Nome del foglio da leggere, se il file è XLSX.

    .PARAMETER IdentityColumn
    Nome esplicito della colonna da usare. Se omesso, viene tentato il riconoscimento automatico.

    .OUTPUTS
    PSCustomObject

    .EXAMPLE
    Import-TeamsUserIdentityList -Path 'C:\Temp\users.xlsx'

    .EXAMPLE
    Import-TeamsUserIdentityList -Path 'C:\Temp\users.csv' -IdentityColumn 'UserPrincipalName'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string] $Path,

        [Parameter()]
        [string] $WorksheetName,

        [Parameter()]
        [string] $IdentityColumn
    )

    $rows = Import-SpreadsheetSafe -Path $Path -WorksheetName $WorksheetName
    if (-not $rows) {
        return @()
    }

    $candidateColumns = @(
        'UserPrincipalName', 'UPN', 'ObjectId', 'Id', 'MailNickName', 'MailNickname', 'Identity', 'User'
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

        [pscustomobject][ordered]@{
            RowNumber          = $rowNumber
            SourceIdentity     = $value
            SourceIdentityType = $identityType
        }
    }
}