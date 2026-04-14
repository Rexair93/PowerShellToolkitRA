@{
    RootModule            = 'CloudOperations.psm1'
    ModuleVersion         = '1.0.0'
    GUID                  = 'ccdb5b64-b6ca-4bc7-ba19-16e1422d7bab' # sostituisci con :NewGuid()
    Author                = 'Renato'
    RequiredModules = @(
        'CloudConnect',
        'DnsTools',
        'FilesUtilities'
    )
    Description           = 'Funzioni di operazioni nel Cloud Microsoft.'
    PowerShellVersion     = '7.0'
    CompatiblePSEditions  = @('Core')
}