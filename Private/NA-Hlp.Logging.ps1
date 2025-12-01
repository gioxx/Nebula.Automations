# Nebula.Automations: (Helpers) Logging =============================================================================================================

function Write-NALog {
    param(
        [string]$Message,
        [ValidateSet('INFO', 'SUCCESS', 'WARNING', 'DEBUG', 'ERROR')]
        [string]$Level = 'INFO',
        [string]$LogLocation
    )

    # Use Write-Log only if it comes from Nebula.Log; otherwise fall back to transcript
    $writeLogCommand = Get-Command -Name Write-Log -ErrorAction SilentlyContinue
    $nebulaWriteLog  = $null
    if ($writeLogCommand) {
        $resolved = if ($writeLogCommand.ResolvedCommand) { $writeLogCommand.ResolvedCommand } else { $writeLogCommand }
        if ($resolved.ModuleName -eq 'Nebula.Log' -or $resolved.Source -eq 'Nebula.Log') {
            $nebulaWriteLog = $writeLogCommand
        }
    }
    if (-not $nebulaWriteLog) {
        $nebulaLogModule = Get-Module -ListAvailable -Name Nebula.Log
        if ($nebulaLogModule) {
            try {
                Import-Module Nebula.Log -ErrorAction Stop
                $writeLogCommand = Get-Command -Name Write-Log -ErrorAction SilentlyContinue
                $resolved = if ($writeLogCommand.ResolvedCommand) { $writeLogCommand.ResolvedCommand } else { $writeLogCommand }
                if ($resolved.ModuleName -eq 'Nebula.Log' -or $resolved.Source -eq 'Nebula.Log') {
                    $nebulaWriteLog = $writeLogCommand
                }
            } catch {
                $nebulaWriteLog = $null
            }
        }
    }

    if ($nebulaWriteLog) {
        try {
            $writeLogParams = @{
                Message     = $Message
                Level       = $Level
                WriteToFile = $true
            }
            $targetLogDir = if (-not [string]::IsNullOrWhiteSpace($LogLocation)) { $LogLocation } else { (Get-Location).Path }
            try {
                if (-not (Test-Path -LiteralPath $targetLogDir)) {
                    $null = New-Item -ItemType Directory -Path $targetLogDir -Force -ErrorAction Stop
                }
                $writeLogParams.LogLocation = $targetLogDir
            } catch {
                # Fall back to transcript-only if the path is not usable
                $writeLogParams.Remove('WriteToFile') | Out-Null
                $writeLogParams.Remove('LogLocation') | Out-Null
                Write-Output "[Nebula.Automations][WARNING] LogLocation '$targetLogDir' non utilizzabile: $($_.Exception.Message)"
            }
            Write-Log @writeLogParams
            return
        } catch {
            # fall back to console output below
        }
    }

    $prefix = "[Nebula.Automations][$Level]"
    Write-Output "$prefix $Message"
}

function ConvertTo-MaskedSecret {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Secret
    )

    if ([string]::IsNullOrWhiteSpace($Secret)) { return '***' }
    if ($Secret.Length -le 8) { return '***' }
    return "$($Secret.Substring(0,4))***$($Secret.Substring($Secret.Length-4,4))"
}
