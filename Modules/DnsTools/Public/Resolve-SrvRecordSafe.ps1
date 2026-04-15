function Resolve-SrvRecordSafe {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string] $Name,

        [Parameter()]
        [string] $Server

    )

    if (-not (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)) {
        Write-Warning "Resolve-DnsName non disponibile su questa piattaforma. Impossibile risolvere SRV per '$Name'."
        return $null
    }

    $nxDomainPattern = '(?i)DNS name does not exist|NXDOMAIN|non esiste|non esistente|No such host is known|Name or service not known'

    try { 
        $params = @{
            Name        = $Name
            Type        = 'SRV'
            ErrorAction = 'Stop'
        }

        if ($Server) {
            $params.Server = $Server
        }

        Resolve-DnsName @params
    }
    catch {
        $err = $_
        $msg = $err.Exception.Message
        $innerMsg = $err.Exception.InnerException?.Message
        $fqid = $err.FullyQualifiedErrorId
        $typeName = $err.Exception.GetType().FullName

        Write-Verbose ("DNS error type: {0} | FQID: {1} | Message: {2}" -f $typeName, $fqid, $msg)

        if (($msg -match $nxDomainPattern) -or ($innerMsg -match $nxDomainPattern)) {
            Write-Verbose "Record SRV non esistente (NXDOMAIN): $Name"
            return
        }

        Write-Warning "Errore DNS risolvendo '$Name': $msg"
        return
    }
}
