# Dot-source internal modules
. "$PSScriptRoot\modules\SendMail.ps1"
. "$PSScriptRoot\modules\ConnectGraph.ps1"

# Export functions
Export-ModuleMember -Function Send-Mail, CheckMGGraphConnection
