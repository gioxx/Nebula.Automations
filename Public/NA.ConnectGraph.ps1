<#
.SYNOPSIS
    Connect to Microsoft Graph using application credentials.
.DESCRIPTION
    Reuses an existing Microsoft Graph session when possible, otherwise performs
    a client credentials flow and connects with Connect-MgGraph -AccessToken.
    Can optionally install Microsoft.Graph if missing.
.EXAMPLE
    Test-MgGraphConnection -TenantId $tenantId -ClientId $clientId -ClientSecret $clientSecret -LogLocation "C:\Logs\graph.log"
.PARAMETER TenantId
    Azure AD tenant ID (GUID or verified domain).
.PARAMETER ClientId
    Application (client) ID used for client-credential auth.
.PARAMETER ClientSecret
    Client secret associated with the application.
.PARAMETER LogLocation
    Path to a log file; used only when Write-Log is available.
.PARAMETER AutoInstall
    Install Microsoft.Graph automatically if missing.
.PARAMETER ShowInformations
    When set, writes additional diagnostic info (never logs secrets).
.NOTES
    Author: Giovanni Solone
#>

function Test-MgGraphConnection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TenantId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientId,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ClientSecret,

        [string]$LogLocation,
        [bool]$AutoInstall = $false,
        [bool]$ShowInformations = $false
    )

    $status = [pscustomobject]@{
        Success   = $false
        Message   = $null
    }

    # Ensure Microsoft.Graph is available
    $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph | Select-Object -First 1
    if (-not $graphModule) {
        if ($AutoInstall) {
            Write-NALog "Microsoft.Graph not found. Installing (CurrentUser scope)..." -Level INFO -LogLocation $LogLocation
            try {
                Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                $graphModule = Get-Module -ListAvailable -Name Microsoft.Graph | Select-Object -First 1
            } catch {
                $status.Message = "Failed to install Microsoft.Graph: $($_.Exception.Message)"
                Write-NALog $status.Message -Level ERROR -LogLocation $LogLocation
                return $status.Success
            }
        }
        else {
            $status.Message = "Microsoft.Graph module is not available. Install it or use -AutoInstall."
            Write-NALog $status.Message -Level ERROR -LogLocation $LogLocation
            return $status.Success
        }
    }

    if (-not (Get-Module -Name Microsoft.Graph)) {
        try {
            Import-Module Microsoft.Graph -ErrorAction Stop
        } catch {
            $status.Message = "Failed to import Microsoft.Graph: $($_.Exception.Message)"
            Write-NALog $status.Message -Level ERROR -LogLocation $LogLocation
            return $status.Success
        }
    }

    try {
        Get-MgUser -Top 1 -ErrorAction Stop
        $status.Success = $true
    }
    catch {

        $Scope = "https://graph.microsoft.com/.default"
        $TokenEndpoint = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"

        $body = @{
            client_id     = $ClientId
            scope         = $Scope
            client_secret = $ClientSecret
            grant_type    = "client_credentials"
        }

        $tokenResponse = $null
        try {
            $tokenResponse = Invoke-RestMethod -Method Post -Uri $TokenEndpoint -Body $body -ErrorAction Stop
        } catch {
            $status.Message = "Failed to obtain access token: $($_.Exception.Message)"
            Write-NALog $status.Message -Level ERROR -LogLocation $LogLocation
            return $status.Success
        }

        if (-not $tokenResponse.access_token) {
            $status.Message = "Token response did not include access_token."
            Write-NALog $status.Message -Level ERROR -LogLocation $LogLocation
            return $status.Success
        }

        if ($ShowInformations) {
            $sanitizedSecret = ConvertTo-MaskedSecret -Secret $ClientSecret
            Write-NALog "TenantId: $TenantId" -Level DEBUG -LogLocation $LogLocation
            Write-NALog "ClientId: $ClientId" -Level DEBUG -LogLocation $LogLocation
            Write-NALog "ClientSecret (masked): $sanitizedSecret" -Level DEBUG -LogLocation $LogLocation
            Write-NALog "Token endpoint: $TokenEndpoint" -Level DEBUG -LogLocation $LogLocation
        }

        $accessToken = $tokenResponse.access_token | ConvertTo-SecureString -AsPlainText -Force

        try {
            Connect-MgGraph -AccessToken $accessToken -ErrorAction Stop
            $status.Success = $true
            $status.Message = "Connected to Microsoft Graph with application credentials."
            Write-NALog $status.Message -Level INFO -LogLocation $LogLocation
        }
        catch {
            $status.Message = "Cannot connect to Microsoft Graph: $($_.Exception.Message)"
            Write-NALog "Connect-MgGraph exception: $($_.Exception.ToString())" -Level ERROR -LogLocation $LogLocation
            Write-NALog $status.Message -Level ERROR -LogLocation $LogLocation
        }
    }
    
    return $status.Success
}

Set-Alias -Name CheckMGGraphConnection -Value Test-MgGraphConnection -Description "Connect to Microsoft Graph (function)"