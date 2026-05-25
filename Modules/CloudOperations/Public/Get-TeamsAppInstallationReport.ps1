function Get-TeamsAppInstallationReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$TeamsAppExternalId,

        [Parameter()]
        [string]$UserFilter = "accountEnabled eq true and userType eq 'Member'",

        [Parameter()]
        [switch]$IncludeAllUsers,

        [Parameter()]
        [ValidateRange(1, 20)]
        [int]$BatchSize = 20,

        [Parameter()]
        [string[]]$Scopes = @('User.Read.All','TeamsAppInstallation.ReadForUser'),

        [Parameter()]
        [switch]$UseDeviceCode,

        [Parameter()]
        [switch]$AutoInstallModules,

        [Parameter()]
        [switch]$ForceReconnect
    )

    begin {
        function Invoke-GraphGetAllPages {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [string]$Uri
            )

            $results = New-Object System.Collections.Generic.List[object]
            $nextLink = $Uri

            while ($nextLink) {
                Write-Verbose "GET $nextLink"
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -OutputType PSObject -ErrorAction Stop

                if ($response.value) {
                    foreach ($item in $response.value) {
                        [void]$results.Add($item)
                    }
                }

                $nextLink = $response.'@odata.nextLink'
            }

            $results
        }

        function Invoke-GraphBatch {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory)]
                [array]$Requests
            )

            $body = @{ requests = $Requests } | ConvertTo-Json -Depth 10
            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/`$batch" -Body $body -ContentType "application/json" -OutputType PSObject -ErrorAction Stop
        }

        if ($AutoInstallModules) {
            Assert-Module -Name Microsoft.Graph.Authentication -Scope CurrentUser -AutoInstall
            Assert-Module -Name Microsoft.Graph.Users -Scope CurrentUser -AutoInstall
        }

        if ($ForceReconnect -and (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue)) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        }

        if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
            throw "Il modulo Microsoft.Graph non è disponibile."
        }

        $connectParams = @{
            Scopes    = $Scopes
            NoWelcome = $true
        }

        if ($UseDeviceCode) {
            $connectParams.UseDeviceCode = $true
        }

        Write-Verbose "Connessione a Microsoft Graph..."
        Connect-MgGraph @connectParams | Out-Null

        $ctx = Get-MgContext
        if (-not $ctx) {
            throw "Connessione a Microsoft Graph non riuscita."
        }

        Write-Verbose ("Account connesso: {0}" -f $ctx.Account)
        Write-Verbose ("Scopes: {0}" -f ($ctx.Scopes -join ', '))

        $selectClause = "id,userPrincipalName,displayName,mail,department,companyName,accountEnabled"
        $usersUri = if ($IncludeAllUsers) {
            "https://graph.microsoft.com/v1.0/users?`$select=$selectClause&`$top=999"
        }
        else {
            $encodedFilter = [uri]::EscapeDataString($UserFilter)
            "https://graph.microsoft.com/v1.0/users?`$filter=$encodedFilter&`$select=$selectClause&`$top=999"
        }

        Write-Verbose "Recupero utenti..."
        $users = @(Invoke-GraphGetAllPages -Uri $usersUri)

        if ($users.Count -eq 0) {
            Write-Warning "Nessun utente trovato."
            return
        }

        Write-Verbose ("Utenti trovati: {0}" -f $users.Count)
        $reportRows = New-Object System.Collections.Generic.List[object]
    }

    process {
        for ($i = 0; $i -lt $users.Count; $i += $BatchSize) {
            $chunk = @($users[$i..([Math]::Min($i + $BatchSize - 1, $users.Count - 1))])

            Write-Verbose ("Batch utenti {0}-{1} di {2}" -f ($i + 1), ($i + $chunk.Count), $users.Count)

            $requests = @()
            $requestMap = @{}
            $reqId = 0

            foreach ($user in $chunk) {
                $reqId++
                $requestId = [string]$reqId

                $requests += @{
                    id     = $requestId
                    method = "GET"
                    url    = "/users/$($user.id)/teamwork/installedApps?`$expand=teamsApp,teamsAppDefinition&`$filter=teamsApp/externalId eq '$TeamsAppExternalId'"
                }

                $requestMap[$requestId] = $user
            }

            $batchResponse = Invoke-GraphBatch -Requests $requests

            foreach ($response in $batchResponse.responses) {
                $user = $requestMap[$response.id]

                $installed = $false
                $installCount = 0
                $errorMessage = $null

                if ($response.status -ge 200 -and $response.status -lt 300) {
                    $items = @($response.body.value)
                    $installCount = $items.Count
                    $installed = $installCount -gt 0
                }
                else {
                    $errorMessage = $response.body.error.message
                }

                [void]$reportRows.Add([pscustomobject][ordered]@{
                    UserId             = $user.id
                    UserPrincipalName  = $user.userPrincipalName
                    DisplayName        = $user.displayName
                    Mail               = $user.mail
                    Department         = $user.department
                    CompanyName        = $user.companyName
                    AccountEnabled     = $user.accountEnabled
                    TeamsAppExternalId = $TeamsAppExternalId
                    IsInstalled        = $installed
                    InstallCount       = $installCount
                    QueryStatus        = $response.status
                    ErrorMessage       = $errorMessage
                })
            }

            Start-Sleep -Milliseconds 200
        }
    }

    end {
        $reportRows
    }
}