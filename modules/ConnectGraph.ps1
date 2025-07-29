<#
.SYNOPSIS
    Connect to Microsoft Graph using the Microsoft.Graph PowerShell module.
.DESCRIPTION
    This function connects to Microsoft Graph using the Microsoft.Graph PowerShell module.
    If the module is not installed, it will attempt to install it.
    If the module is installed but not loaded, it will attempt to load it.
    If the module is loaded but the connection fails, it will attempt to connect using the provided credentials.
    If the connection is successful, it will return $true.
    If the connection fails, it will return $false.
.PARAMETER tenantId
    The tenant ID of the Azure AD tenant.
.PARAMETER clientId
    The client ID of the Azure AD application.
.PARAMETER clientSecret
    The client secret of the Azure AD application.
.PARAMETER LogLocation
    The location where log files will be stored.
.PARAMETER showInformations
    If set to $true, the function will log additional information to the log file.
.EXAMPLE
    CheckMGGraphConnection -tenantId $tenantId -clientId $clientId -clientSecret $clientSecret -LogLocation $LogLocation

    This example connects to Microsoft Graph using the provided tenant ID, client ID, and client secret.
.NOTES
    Author: Giovanni Solone

    Modification History:
    - 2025-03-28: Initial version (isolation of Microsoft Graph check / connection function from main script).
#>

#Requires -Module Microsoft.Graph
#Requires -Module Nebula.Log

# Connect to Microsoft Graph Function
function CheckMGGraphConnection {
    param (
        [string] $tenantId,
        [string] $clientId,
        [string] $clientSecret,
        [string] $LogLocation,
        [bool] $showInformations = $false
    )

    $mggConnected = $false

    if ( (Get-Module -Name Microsoft.Graph -ListAvailable).count -gt 0 ) {
        try {
            Get-MgUser -ErrorAction Stop
            $mggConnected = $true
        } catch {
            Write-Log -LogLocation $LogLocation -Message "Please wait until I load Microsoft Graph, the operation may take a minute or more ..." -Level "INFO" -WriteToFile

            $scope = "https://graph.microsoft.com/.default"
            $tokenEndpoint = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

            $body = @{
                client_id     = $clientId
                scope         = $scope
                client_secret = $clientSecret
                grant_type    = "client_credentials"
            }

            $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenEndpoint -Body $body
            $accessToken = $tokenResponse.access_token | ConvertTo-SecureString -AsPlainText -Force
            
            if ($showInformations) {
                Write-Log -LogLocation $LogLocation -Message "Tenant ID: $tenantId" -Level "DEBUG"
                Write-Log -LogLocation $LogLocation -Message "Client ID: $clientId" -Level "DEBUG"
                Write-Log -LogLocation $LogLocation -Message "Client Secret: $clientSecret" -Level "DEBUG"
                Write-Log -LogLocation $LogLocation -Message "Scope: $scope" -Level "DEBUG"
                Write-Log -LogLocation $LogLocation -Message "Token Endpoint: $tokenEndpoint" -Level "DEBUG"
                Write-Log -LogLocation $LogLocation -Message "Access Token: $accessToken" -Level "DEBUG"
            }

            try {
                Connect-MgGraph -AccessToken $accessToken -ErrorAction Stop
                $mggConnected = $true
            } catch {
                Write-Log -LogLocation $LogLocation -Message "Cannot connect to Microsoft Graph. Error: $($_.Exception.Message)" -Level "ERROR" -WriteToFile
                $mggConnected = $false
            }
            
        }
    }

    return $mggConnected
}