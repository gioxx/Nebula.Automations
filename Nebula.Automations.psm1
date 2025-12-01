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
    Set-Alias -Name Log-Message -Value Write-Log -ErrorAction SilentlyContinue
}
