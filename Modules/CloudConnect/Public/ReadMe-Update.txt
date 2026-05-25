Come usarla

Per aggiornare tutto e pulire le versioni vecchie:
powershell
Update-InstalledModules -CleanupOldVersions -Verbose

Per testare senza modificare nulla:
powershell
Update-InstalledModules -CleanupOldVersions -WhatIf -Verbose

Per limitarti ad alcuni moduli:
powershell
Update-InstalledModules -Name Microsoft.Graph.Authentication,MicrosoftTeams -CleanupOldVersions -Verbose

L’uso di Uninstall-Module -RequiredVersion è la parte corretta per rimuovere solo le release precedenti, perché PowerShellGet gestisce la disinstallazione per versione specifica e non elimina in automatico le versioni obsolete dopo Update-Module.