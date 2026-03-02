# Nebula.Automations.psm1
$script:ModuleRoot = $PSScriptRoot

# --- Load Private helpers first (NOT exported) ---
$privateDir = Join-Path $PSScriptRoot 'Private'
if (Test-Path $privateDir) {
    Get-ChildItem -Path $privateDir -Filter '*.ps1' -File -Recurse | Sort-Object FullName | ForEach-Object {
        try {
            . $_.FullName  # dot-source
        }
        catch {
            throw "Failed to load Private script '$($_.Name)': $($_.Exception.Message)"
        }
    }
}

# --- Load Public entry points (will be exported) ---
$publicDir = Join-Path $PSScriptRoot 'Public'
if (Test-Path $publicDir) {
    Get-ChildItem -Path $publicDir -Filter '*.ps1' -File | ForEach-Object {
        try {
            . $_.FullName  # dot-source
        }
        catch {
            throw "Failed to load Public script '$($_.Name)': $($_.Exception.Message)"
        }
    }
}

# Fallback for Write-Log (if Nebula.Log is not installed on the system)
if (-not (Get-Command -Name Write-Log -ErrorAction SilentlyContinue)) {
    function Write-Log {
        <#
        .SYNOPSIS
            Compatibility logging function exposed when Nebula.Log is not installed.
        .DESCRIPTION
            Provides a Write-Log-compatible surface and delegates logging to Write-NALog.
            Useful for scripts that already call Write-Log/Log-Message.
        .PARAMETER Message
            Log message text.
        .PARAMETER Level
            Log severity: INFO, SUCCESS, WARNING, DEBUG, ERROR.
        .PARAMETER LogLocation
            Optional path to a log file.
        .PARAMETER WriteToFile
            Compatibility switch kept for signature parity; ignored in fallback mode.
        .EXAMPLE
            Write-Log -Message 'Hello from Nebula.Automations' -Level INFO
        .EXAMPLE
            Write-Log -Message 'Task completed' -Level SUCCESS -LogLocation 'C:\Logs\MyScript'
        .LINK
            https://kb.gioxx.org/Nebula/Automations/usage/write-log
        #>
        [CmdletBinding()]
        param(
            [string]$Message,
            [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'DEBUG', 'ERROR')]
            [string]$Level = 'INFO',
            [string]$LogLocation,
            [switch]$WriteToFile
        )
        # delegate to Write-NALog; ignore WriteToFile if Nebula.Log is missing
        Write-NALog -Message $Message -Level $Level -LogLocation $LogLocation
    }
}

# --- Aliases & Exports -------------------------------------------------------
$existingLogAlias = Get-Alias -Name 'Log-Message' -ErrorAction SilentlyContinue
if (-not $existingLogAlias -or $existingLogAlias.ResolvedCommandName -ne 'Write-Log') {
    Set-Alias -Name Log-Message -Value Write-Log -Force
}

$existingGraphAlias = Get-Alias -Name 'CheckMGGraphConnection' -ErrorAction SilentlyContinue
if (-not $existingGraphAlias -or $existingGraphAlias.ResolvedCommandName -ne 'Test-MgGraphConnection') {
    Set-Alias -Name CheckMGGraphConnection -Value Test-MgGraphConnection -Force
}

