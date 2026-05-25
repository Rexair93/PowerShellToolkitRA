function Update-InstalledModules {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter()]
        [string[]] $Name,

        [Parameter()]
        [ValidateSet('CurrentUser', 'AllUsers')]
        [string] $Scope = 'CurrentUser',

        [Parameter()]
        [switch] $AllowClobber,

        [Parameter()]
        [switch] $Force,

        [Parameter()]
        [switch] $CleanupOldVersions,

        [Parameter()]
        [switch] $IncludePrerelease
    )

    if (-not (Get-Command Get-InstalledModule -ErrorAction SilentlyContinue)) {
        throw "Get-InstalledModule non disponibile. Installa o aggiorna PowerShellGet."
    }
    if (-not (Get-Command Find-Module -ErrorAction SilentlyContinue)) {
        throw "Find-Module non disponibile. Installa o aggiorna PowerShellGet."
    }
    if (-not (Get-Command Update-Module -ErrorAction SilentlyContinue)) {
        throw "Update-Module non disponibile. Installa o aggiorna PowerShellGet."
    }
    if ($CleanupOldVersions -and -not (Get-Command Uninstall-Module -ErrorAction SilentlyContinue)) {
        throw "Uninstall-Module non disponibile. Installa o aggiorna PowerShellGet."
    }

    # Ottieni i nomi da processare: se forniti, usa quelli; altrimenti ricava i nomi unici dei moduli installati
    if ($Name -and $Name.Count -gt 0) {
        $namesToProcess = $Name | Sort-Object -Unique
    }
    else {
        $namesToProcess = Get-InstalledModule -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty Name |
            Sort-Object -Unique
    }

    if (-not $namesToProcess -or $namesToProcess.Count -eq 0) {
        Write-Verbose "Nessun modulo da processare."
        return
    }

    foreach ($moduleName in $namesToProcess) {
        Write-Verbose "Elaborazione modulo: $moduleName"

        # Leggi tutte le versioni installate per questo nome (singola chiamata con -Name)
        try {
            $installedVersions = Get-InstalledModule -Name $moduleName -AllVersions -ErrorAction Stop |
                Sort-Object Version -Descending
        }
        catch {
            Write-Warning "Impossibile leggere le versioni installate di '$moduleName': $($_.Exception.Message)"
            continue
        }

        if (-not $installedVersions -or $installedVersions.Count -eq 0) {
            Write-Verbose "Nessuna versione installata trovata per $moduleName"
            continue
        }

        $latestInstalled = $installedVersions | Select-Object -First 1

        # Verifica repository per versione più recente
        $findParams = @{ Name = $moduleName; ErrorAction = 'Stop' }
        if ($IncludePrerelease) { $findParams.AllowPrereleaseVersions = $true }

        try {
            $galleryModule = Find-Module @findParams
        }
        catch {
            Write-Warning "Impossibile verificare aggiornamenti per '$moduleName': $($_.Exception.Message)"
            continue
        }

        if (-not $galleryModule) {
            Write-Verbose "Modulo '$moduleName' non presente in repository pubblico/compatibile."
            continue
        }

        $hasUpdate = [version]$galleryModule.Version -gt [version]$latestInstalled.Version

        if ($hasUpdate) {
            $updateParams = @{
                Name        = $moduleName
                ErrorAction = 'Stop'
            }
            if ($AllowClobber) { $updateParams.AllowClobber = $true }
            if ($Force) { $updateParams.Force = $true }
            if ($Scope) { $updateParams.Scope = $Scope }
            if ($IncludePrerelease) { $updateParams.AllowPrereleaseVersions = $true }

            if ($PSCmdlet.ShouldProcess($moduleName, "Aggiornare dalla versione $($latestInstalled.Version) alla versione $($galleryModule.Version)")) {
                try {
                    Write-Verbose "Aggiornamento di '$moduleName' alla versione $($galleryModule.Version)..."
                    Update-Module @updateParams
                }
                catch {
                    Write-Warning "Aggiornamento fallito per '$moduleName': $($_.Exception.Message)"
                    # prosegui comunque con cleanup se richiesto (alcuni moduli possono avere più installazioni)
                }
            }
        }
        else {
            Write-Verbose "Nessun aggiornamento disponibile per '$moduleName' (installato: $($latestInstalled.Version))."
        }

        if ($CleanupOldVersions) {
            try {
                $installedAfter = Get-InstalledModule -Name $moduleName -AllVersions -ErrorAction Stop |
                    Sort-Object Version -Descending
            }
            catch {
                Write-Warning "Impossibile rileggere le versioni installate di '$moduleName' dopo l'aggiornamento: $($_.Exception.Message)"
                continue
            }

            $toKeep = $installedAfter | Select-Object -First 1
            $toRemove = $installedAfter | Select-Object -Skip 1

            foreach ($old in $toRemove) {
                if ($PSCmdlet.ShouldProcess($moduleName, "Rimuovere la vecchia versione $($old.Version)")) {
                    try {
                        Write-Verbose "Rimozione di '$moduleName' versione $($old.Version)..."
                        Uninstall-Module -Name $moduleName -RequiredVersion $old.Version -Force -ErrorAction Stop
                    }
                    catch {
                        Write-Warning "Fallita rimozione di '$moduleName' versione $($old.Version): $($_.Exception.Message)"
                    }
                }
            }

            if ($toKeep) {
                Write-Verbose "Mantenuta versione $($toKeep.Version) per '$moduleName'."
            }
        }
    }
}