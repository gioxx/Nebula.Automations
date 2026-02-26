# Nebula.Automations: (Helpers) Security ============================================================================================================

function ConvertTo-MaskedSecret {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Secret
    )

    if ([string]::IsNullOrWhiteSpace($Secret)) { return '***' }
    if ($Secret.Length -le 8) { return '***' }
    return "$($Secret.Substring(0,4))***$($Secret.Substring($Secret.Length-4,4))"
}
