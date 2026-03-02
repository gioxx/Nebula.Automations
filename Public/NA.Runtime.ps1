function Import-PreferredModule {
    <#
    .SYNOPSIS
        Imports a module preferring a development manifest when available.
    .DESCRIPTION
        Tries to import a module from DevManifestPath first (if requested and found),
        otherwise imports the installed module by name.
    .PARAMETER ModuleName
        Module name to import.
    .PARAMETER DevManifestPath
        Optional path to the development module manifest (*.psd1).
    .PARAMETER PreferDev
        Prefer DevManifestPath when it exists. Enabled by default.
    .PARAMETER Force
        Force module reload.
    .EXAMPLE
        Import-PreferredModule -ModuleName 'Nebula.Automations' -DevManifestPath 'C:\Temp\Nebula.Automations\Nebula.Automations.psd1'
    .EXAMPLE
        Import-PreferredModule -ModuleName 'Nebula.Log' -PreferDev $false -Force
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/import-preferredmodule
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ModuleName,

        [string]$DevManifestPath,

        [bool]$PreferDev = $true,

        [switch]$Force
    )

    $result = [pscustomobject]@{
        Success    = $false
        ModuleName = $ModuleName
        Source     = 'Installed'
        Version    = $null
        Path       = $null
        Message    = $null
    }

    try {
        if ($Force.IsPresent) {
            Remove-Module $ModuleName -Force -ErrorAction SilentlyContinue
        }

        $useDev = $PreferDev -and -not [string]::IsNullOrWhiteSpace($DevManifestPath) -and (Test-Path -LiteralPath $DevManifestPath)
        $resolvedDevManifestPath = $null
        if ($useDev) {
            $resolvedDevManifestPath = (Resolve-Path -LiteralPath $DevManifestPath -ErrorAction Stop).Path
            Import-Module $DevManifestPath -Force:$Force.IsPresent -WarningAction SilentlyContinue -ErrorAction Stop
        } else {
            Import-Module $ModuleName -Force:$Force.IsPresent -WarningAction SilentlyContinue -ErrorAction Stop
        }

        $loadedModule = Get-Module -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $loadedModule) {
            throw "Module '$ModuleName' appears not loaded after import attempt."
        }

        $loadedModuleDir = Split-Path -Path $loadedModule.Path -Parent
        if (-not [string]::IsNullOrWhiteSpace($resolvedDevManifestPath)) {
            $devModuleDir = Split-Path -Path $resolvedDevManifestPath -Parent
            if ([string]::Equals($loadedModuleDir, $devModuleDir, [System.StringComparison]::OrdinalIgnoreCase)) {
                $result.Source = 'DEV'
            } else {
                $result.Source = 'Installed'
            }
        } else {
            $result.Source = 'Installed'
        }

        $result.Success = $true
        $result.Version = $loadedModule.Version
        $result.Path = $loadedModule.Path
        $result.Message = "Module '$ModuleName' loaded ($($result.Source)): v$($result.Version) from $($result.Path)"
    } catch {
        $result.Message = $_.Exception.Message
    }

    return $result
}


function Initialize-ScriptRuntime {
    <#
    .SYNOPSIS
        Initializes common script runtime prerequisites.
    .DESCRIPTION
        Imports requested modules, loads an XML configuration file and optionally
        ensures that a log directory exists.
    .PARAMETER ConfigPath
        Full path to the XML configuration file.
    .PARAMETER ModulesToImport
        Modules to import before script execution.
    .PARAMETER LogDirectory
        Log directory path.
    .PARAMETER EnsureLogDirectory
        Create LogDirectory if it does not exist.
    .EXAMPLE
        Initialize-ScriptRuntime -ConfigPath 'C:\Config\tenant.config.xml' -ModulesToImport @('Nebula.Log','Nebula.Automations') -LogDirectory 'C:\Logs' -EnsureLogDirectory
    .EXAMPLE
        $runtime = Initialize-ScriptRuntime -ConfigPath 'C:\Config\tenant.config.xml'
        if (-not $runtime.Success) { throw $runtime.Message }
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/initialize-scriptruntime
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ConfigPath,

        [string[]]$ModulesToImport = @('Nebula.Log', 'Nebula.Automations'),

        [string]$LogDirectory,

        [switch]$EnsureLogDirectory
    )

    $result = [pscustomobject]@{
        Success      = $false
        Config       = $null
        ConfigPath   = $ConfigPath
        LogDirectory = $LogDirectory
        Message      = $null
    }

    try {
        foreach ($moduleName in $ModulesToImport) {
            if ([string]::IsNullOrWhiteSpace($moduleName)) { continue }
            Import-Module $moduleName -Force -WarningAction SilentlyContinue -ErrorAction Stop
        }

        if (-not (Test-Path -LiteralPath $ConfigPath)) {
            throw "Configuration file not found: $ConfigPath"
        }

        [xml]$config = Get-Content -LiteralPath $ConfigPath -ErrorAction Stop
        $result.Config = $config

        if ($EnsureLogDirectory.IsPresent -and -not [string]::IsNullOrWhiteSpace($LogDirectory)) {
            if (-not (Test-Path -LiteralPath $LogDirectory)) {
                $null = New-Item -ItemType Directory -Path $LogDirectory -Force -ErrorAction Stop
            }
        }

        $result.Success = $true
        $result.Message = 'Runtime initialization completed.'
    } catch {
        $result.Message = $_.Exception.Message
    }

    return $result
}

function Resolve-ScriptConfigPaths {
    <#
    .SYNOPSIS
        Builds common script paths (config/log/output) from a script root.
    .DESCRIPTION
        Resolves a reusable path set used by automation scripts:
        - Config root + config file path
        - Log directory path
        - Output file path
    .PARAMETER ScriptRoot
        Root folder of the calling script (typically $PSScriptRoot).
    .PARAMETER ConfigRelativePath
        Relative path of the configuration file under ConfigRoot.
    .PARAMETER ConfigRootPath
        Optional explicit config root. If omitted, parent of ScriptRoot is used.
    .PARAMETER LogRelativePath
        Optional relative log directory under ScriptRoot.
    .PARAMETER OutputRelativePath
        Optional relative output file path under ScriptRoot.
    .PARAMETER EnsureDirectories
        Creates log directory and output parent directory if missing.
    .EXAMPLE
        Resolve-ScriptConfigPaths -ScriptRoot $PSScriptRoot -ConfigRelativePath 'Config\tenant.config.xml' -LogRelativePath 'Logs\MyScript' -OutputRelativePath 'Export\result.json'
    .EXAMPLE
        Resolve-ScriptConfigPaths -ScriptRoot $PSScriptRoot -ConfigRootPath 'C:\AutomationRoot' -ConfigRelativePath 'Config\tenant.config.xml' -EnsureDirectories
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/resolve-scriptconfigpaths
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ScriptRoot,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ConfigRelativePath,

        [string]$ConfigRootPath,
        [string]$LogRelativePath,
        [string]$OutputRelativePath,
        [switch]$EnsureDirectories
    )

    $result = [pscustomobject]@{
        Success      = $false
        ScriptRoot   = $null
        ConfigRoot   = $null
        ConfigPath   = $null
        LogDirectory = $null
        OutputPath   = $null
        Message      = $null
    }

    try {
        if (-not (Test-Path -LiteralPath $ScriptRoot)) {
            throw "ScriptRoot not found: $ScriptRoot"
        }

        $resolvedScriptRoot = (Resolve-Path -LiteralPath $ScriptRoot -ErrorAction Stop).Path
        $resolvedConfigRoot = if (-not [string]::IsNullOrWhiteSpace($ConfigRootPath)) {
            if (-not (Test-Path -LiteralPath $ConfigRootPath)) {
                throw "ConfigRootPath not found: $ConfigRootPath"
            }
            (Resolve-Path -LiteralPath $ConfigRootPath -ErrorAction Stop).Path
        } else {
            Split-Path -Path $resolvedScriptRoot -Parent
        }

        $result.ScriptRoot = $resolvedScriptRoot
        $result.ConfigRoot = $resolvedConfigRoot
        $result.ConfigPath = Join-Path -Path $resolvedConfigRoot -ChildPath $ConfigRelativePath

        if (-not [string]::IsNullOrWhiteSpace($LogRelativePath)) {
            $result.LogDirectory = Join-Path -Path $resolvedScriptRoot -ChildPath $LogRelativePath
        }
        if (-not [string]::IsNullOrWhiteSpace($OutputRelativePath)) {
            $result.OutputPath = Join-Path -Path $resolvedScriptRoot -ChildPath $OutputRelativePath
        }

        if ($EnsureDirectories.IsPresent) {
            if (-not [string]::IsNullOrWhiteSpace($result.LogDirectory) -and -not (Test-Path -LiteralPath $result.LogDirectory)) {
                $null = New-Item -ItemType Directory -Path $result.LogDirectory -Force -ErrorAction Stop
            }

            if (-not [string]::IsNullOrWhiteSpace($result.OutputPath)) {
                $outputParent = Split-Path -Path $result.OutputPath -Parent
                if (-not [string]::IsNullOrWhiteSpace($outputParent) -and -not (Test-Path -LiteralPath $outputParent)) {
                    $null = New-Item -ItemType Directory -Path $outputParent -Force -ErrorAction Stop
                }
            }
        }

        $result.Success = $true
        $result.Message = 'Path resolution completed.'
    } catch {
        $result.Message = $_.Exception.Message
    }

    return $result
}


function Test-ScriptActivityLog {
    <#
    .SYNOPSIS
        Verifies that activity logging is available for a script.
    .DESCRIPTION
        Uses Test-ActivityLog when available (typically from Nebula.Log).
        Falls back to a direct write/delete test in the log directory.
    .PARAMETER LogLocation
        Log directory path to validate.
    .EXAMPLE
        Test-ScriptActivityLog -LogLocation 'C:\Logs\MyScript'
    .EXAMPLE
        $status = Test-ScriptActivityLog -LogLocation 'C:\Logs\MyScript'
        if ($status.Status -ne 'OK') { throw $status.Message }
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/test-scriptactivitylog
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$LogLocation
    )

    $result = [pscustomobject]@{
        Success     = $false
        Status      = 'KO'
        LogLocation = $LogLocation
        Message     = $null
    }

    try {
        if (-not (Test-Path -LiteralPath $LogLocation)) {
            $null = New-Item -ItemType Directory -Path $LogLocation -Force -ErrorAction Stop
        }

        $testActivityLogCommand = Get-Command -Name 'Test-ActivityLog' -ErrorAction SilentlyContinue
        if ($testActivityLogCommand) {
            $status = Test-ActivityLog -LogLocation $LogLocation
            if ($status -eq 'OK') {
                $result.Success = $true
                $result.Status = 'OK'
                $result.Message = 'Activity log is ready and writable.'
            } else {
                $result.Message = "Test-ActivityLog returned '$status'."
            }
        } else {
            $testFile = Join-Path -Path $LogLocation -ChildPath '.na-activitylog.test'
            Set-Content -Path $testFile -Value 'ok' -Encoding UTF8 -ErrorAction Stop
            Remove-Item -Path $testFile -Force -ErrorAction Stop
            $result.Success = $true
            $result.Status = 'OK'
            $result.Message = 'Activity log fallback test passed.'
        }

        if (Get-Command -Name 'Write-NALog' -ErrorAction SilentlyContinue) {
            $level = if ($result.Success) { 'INFO' } else { 'ERROR' }
            Write-NALog -Message $result.Message -Level $level -LogLocation $LogLocation
        }
    } catch {
        $result.Message = $_.Exception.Message
        if (Get-Command -Name 'Write-NALog' -ErrorAction SilentlyContinue) {
            Write-NALog -Message "Activity log check failed: $($result.Message)" -Level ERROR -LogLocation $LogLocation
        }
    }

    return $result
}


function Start-ScriptTranscript {
    <#
    .SYNOPSIS
        Starts a transcript safely for automation scripts.
    .DESCRIPTION
        Optionally clears old transcript files and starts a new transcript in the
        specified output directory.
    .PARAMETER OutputDirectory
        Output directory where transcript files are written.
    .PARAMETER CleanupOld
        Remove older transcript files matching CleanupPattern.
    .PARAMETER CleanupPattern
        Pattern used when CleanupOld is enabled.
    .PARAMETER IncludeInvocationHeader
        Include invocation header in the transcript.
    .EXAMPLE
        Start-ScriptTranscript -OutputDirectory 'C:\Logs\MyScript' -CleanupOld
    .EXAMPLE
        Start-ScriptTranscript -OutputDirectory 'C:\Logs\MyScript' -IncludeInvocationHeader
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/start-scripttranscript
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputDirectory,

        [switch]$CleanupOld,

        [string]$CleanupPattern = 'PowerShell*.txt',

        [switch]$IncludeInvocationHeader
    )

    $result = [pscustomobject]@{
        Success        = $false
        TranscriptPath = $null
        Message        = $null
    }

    try {
        if (-not (Test-Path -LiteralPath $OutputDirectory)) {
            $null = New-Item -ItemType Directory -Path $OutputDirectory -Force -ErrorAction Stop
        }

        if ($CleanupOld.IsPresent -and -not [string]::IsNullOrWhiteSpace($CleanupPattern)) {
            Get-ChildItem -Path $OutputDirectory -Filter $CleanupPattern -File -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }

        $startParams = @{
            OutputDirectory = $OutputDirectory
        }
        if ($IncludeInvocationHeader.IsPresent) {
            $startParams.IncludeInvocationHeader = $true
        }

        $transcript = Start-Transcript @startParams
        $result.TranscriptPath = $transcript.Path
        $result.Success = $true
        $result.Message = 'Transcript started successfully.'
    } catch {
        $result.Message = $_.Exception.Message
    }

    return $result
}


function Stop-ScriptTranscriptSafe {
    <#
    .SYNOPSIS
        Stops transcript safely without throwing if no transcript is active.
    .DESCRIPTION
        Attempts to stop transcript and returns a status object.
    .EXAMPLE
        Stop-ScriptTranscriptSafe
    .EXAMPLE
        $stop = Stop-ScriptTranscriptSafe
        Write-Output $stop.Message
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/stop-scripttranscriptsafe
    #>
    [CmdletBinding()]
    param()

    $result = [pscustomobject]@{
        Success = $false
        Message = $null
    }

    try {
        Stop-Transcript -ErrorAction Stop | Out-Null
        $result.Success = $true
        $result.Message = 'Transcript stopped successfully.'
    } catch {
        if ($_.Exception.Message -match 'transcription has not been started') {
            $result.Success = $true
            $result.Message = 'No active transcript. Nothing to stop.'
        } else {
            $result.Message = $_.Exception.Message
        }
    }

    return $result
}
