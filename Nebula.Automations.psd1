@{
    RootModule        = 'Nebula.Automations.psm1'
    ModuleVersion     = '1.2.0'
    GUID              = 'b94d3242-e96d-4078-ab12-c31a3f0221c2'
    Author            = 'Giovanni Solone'
    Description       = 'Common utilities for PowerShell scripting and automations: mail, Graph connectivity, scheduled tasks, and more.'

    # Minimum required PowerShell (PS 5.1 works; better with PS 7+)
    PowerShellVersion    = '5.1'
    CompatiblePSEditions = @('Desktop', 'Core')
    RequiredAssemblies   = @()
    FunctionsToExport = @(
        'Import-PreferredModule',
        'Initialize-ScriptRuntime',
        'Invoke-ScriptTaskLifecycle',
        'Register-ScriptScheduledTask',
        'Resolve-ScriptConfigPaths',
        'Send-Mail',
        'Send-ReportIfChanged',
        'Start-ScriptTranscript',
        'Stop-ScriptTranscriptSafe',
        'Test-MgGraphConnection',
        'Test-ScriptActivityLog',
        'Unregister-ScriptScheduledTask',
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
            Tags         = @('Microsoft', 'Automations', 'PowerShell', 'Graph', 'Mail', 'ScheduledTask', 'TaskScheduler', 'Nebula', 'Utilities')
            ProjectUri   = 'https://github.com/gioxx/Nebula.Automations'
            LicenseUri   = 'https://opensource.org/licenses/MIT'
            IconUri      = 'https://raw.githubusercontent.com/gioxx/Nebula.Automations/main/Assets/icon.png'
            ReleaseNotes = @'
- Improve: Comment-based help links updated for scheduled task usage documentation.
- Improve: Internal helpers refined in Private scope (logging/security split).
- Improve: Module aliases are now centralized in Nebula.Automations.psm1 for consistency.
- Improve: Public scripts reorganized by domain (Connections, Messaging, ScheduledTasks).
- New: Import-PreferredModule helper to prefer DEV module manifests with fallback to installed modules.
- New: Initialize-ScriptRuntime to centralize module import, config loading and log directory bootstrap.
- New: Invoke-ScriptTaskLifecycle function to delegate register/unregister flow with optional credential prompt and HH:mm parsing.
- New: Register-ScriptScheduledTask and Unregister-ScriptScheduledTask functions.
- New: Resolve-ScriptConfigPaths helper to standardize config/log/output path discovery in scripts.
- New: Send-ReportIfChanged helper to centralize conditional report dispatch when changes are detected.
- New: Start-ScriptTranscript and Stop-ScriptTranscriptSafe helpers for reusable transcript lifecycle handling.
- New: Test-ScriptActivityLog helper to centralize activity-log readiness checks.
'@
        }
    }
}
