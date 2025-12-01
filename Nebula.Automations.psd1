@{
    RootModule        = 'Nebula.Automations.psm1'
    ModuleVersion     = '1.1.0'
    GUID              = 'b94d3242-e96d-4078-ab12-c31a3f0221c2'
    Author            = 'Giovanni Solone'
    Description       = 'Common utilities for PowerShell scripting and automations: mail, Graph connectivity, and more.'

    # Minimum required PowerShell (PS 5.1 works; better with PS 7+)
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    RequiredAssemblies   = @()
    FunctionsToExport = @(
        'Send-Mail',
        'Test-MgGraphConnection',
        'Write-Log'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @(
        'CheckMGGraphConnection',
        'Log-Message'
    )

    PrivateData       = @{
        PSData = @{
            Tags         = @('Microsoft', 'Automations', 'PowerShell', 'Graph', 'Mail', 'Nebula', 'Utilities')
            ProjectUri   = 'https://github.com/gioxx/Nebula.Automations'
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Automations/main/Assets/icon.png'
            ReleaseNotes = @'
- The module has been completely refactored to improve maintainability and future extensibility.
- Logging: added Write-Log/Log-Message fallback exported by the module when Nebula.Log is missing; Write-NALog now calls Write-Log only when it belongs to Nebula.Log.
- Graph: Test-MgGraphConnection now performs client-credential flow and Connect-MgGraph with an access token string; parameters AutoInstall/ShowInformations are boolean, alias CheckMGGraphConnection retained.
'@
        }
    }
}
