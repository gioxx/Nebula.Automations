@{
    RootModule        = 'Nebula.Automations.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'b94d3242-e96d-4078-ab12-c31a3f0221c2'
    Author            = 'Giovanni Solone'
    Description       = 'Common utilities for PowerShell scripting and automations: mail, Graph connectivity, and more.'

    PowerShellVersion = '7.0'

    FunctionsToExport = @(
        'Send-Mail',
        'CheckMGGraphConnection'
    )

    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    PrivateData       = @{
        PSData = @{
            Tags         = @('Automations', 'PowerShell', 'Graph', 'Mail', 'Nebula', 'Utilities')
            License      = 'MIT'
            ProjectUri   = 'https://github.com/gioxx/Nebula.Automations'
            Icon         = 'icon.png'
            Readme       = 'README.md'
            ReleaseNotes = @'
- Minor fixes and improvements.
- Module rename: Nebula.Automations (formerly Nebula.Tools). Previous GUID was invalidated (d6f6c63d-e8db-4f0c-b7f6-4b0a95f7a63e).
'@
        }
    }
}
