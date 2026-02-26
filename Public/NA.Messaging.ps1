<#
.SYNOPSIS
    E-mail functions for PowerShell scripts.
.DESCRIPTION
    This module contains functions for sending e-mails.
    The Send-Mail function sends an e-mail using the SMTP server and port specified.
.PARAMETER SMTPServer
    The SMTP server to use for sending the e-mail.
.PARAMETER SMTPPort
    The SMTP port to use for sending the e-mail. Default is 25.
.PARAMETER From
    The e-mail address of the sender.
.PARAMETER To
    The e-mail address of the recipient.
.PARAMETER CC
    The e-mail address of the recipient in CC (Carbon Copy).
.PARAMETER Bcc
    The e-mail address of the recipient in BCC (Blind Carbon Copy).
.PARAMETER Subject
    The subject of the e-mail.
.PARAMETER Body
    The body of the e-mail.
.PARAMETER AttachmentPath
    The path to the file to attach to the e-mail.
.EXAMPLE
    Send-Mail -SMTPServer "smtp.contoso.com" -From "user@contoso.com" -To "sharedmailbox@contoso.com" -Subject "Job completed" -Body "<p>All good.</p>"
.EXAMPLE
    Send-Mail -SMTPServer "smtp.contoso.com" -SMTPPort 587 -UseSsl -Credential (Get-Credential) -From "user@contoso.com" -To "sharedmailbox@contoso.com","user2@contoso.com" -Cc "user3@contoso.com" -Subject "Weekly report" -Body "<p>See attachment.</p>" -AttachmentPath "C:\Reports\weekly.csv"
.LINK
    https://kb.gioxx.org/Nebula/Automations/usage/send-mail
.NOTES
    Author: Giovanni Solone
#>

# Mail Function
function Send-Mail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $SMTPServer,

        [ValidateRange(1, 65535)]
        [int] $SMTPPort = 25,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $From,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]] $To,

        [string[]] $Cc,
        [string[]] $Bcc,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $Subject,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string] $Body,

        [string[]] $AttachmentPath,

        [switch] $PlainText,
        [System.Management.Automation.PSCredential] $Credential,
        [switch] $UseSsl
    )

    try {
        $mailMessage = New-Object System.Net.Mail.MailMessage
        $mailMessage.From = $From
        $To | Where-Object { $_ } | ForEach-Object { [void]$mailMessage.To.Add($_) }
        $Cc | Where-Object { $_ } | ForEach-Object { [void]$mailMessage.CC.Add($_) }
        $Bcc | Where-Object { $_ } | ForEach-Object { [void]$mailMessage.Bcc.Add($_) }

        $mailMessage.Subject = $Subject
        $mailMessage.Body = $Body
        $mailMessage.IsBodyHtml = -not $PlainText.IsPresent  # default HTML to preserve old behavior

        if ($AttachmentPath) {
            foreach ($path in $AttachmentPath) {
                if (-not [string]::IsNullOrWhiteSpace($path)) {
                    if (-not (Test-Path -Path $path)) {
                        throw "Attachment not found: $path"
                    }
                    $attachment = New-Object System.Net.Mail.Attachment($path)
                    $mailMessage.Attachments.Add($attachment) | Out-Null
                }
            }
        }

        $smtpClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
        $smtpClient.EnableSsl = $UseSsl.IsPresent
        if ($Credential) {
            $smtpClient.Credentials = $Credential
        }

        $smtpClient.Send($mailMessage)
        Write-NALog "Mail sent to: $($To -join ', ')" -Level INFO
    } catch {
        Write-NALog "Failed to send e-mail: $($_.Exception.Message)" -Level ERROR
        throw
    } finally {
        if ($mailMessage -and $mailMessage.Attachments.Count -gt 0) {
            $mailMessage.Attachments.Dispose()
        }
        if ($mailMessage) {
            $mailMessage.Dispose()
        }
    }
}


function Send-ReportIfChanged {
    <#
    .SYNOPSIS
        Sends a report e-mail only when changes are present.
    .DESCRIPTION
        Finalizes an HTML report body and sends it only if:
        - SendLogs is true
        - ModCounter is greater than 0
    .PARAMETER SendLogs
        Enables or disables report sending.
    .PARAMETER ModCounter
        Number of detected changes.
    .PARAMETER MailBody
        HTML body prefix/content to finalize before sending.
    .PARAMETER SmtpServer
        SMTP server.
    .PARAMETER From
        Sender e-mail address.
    .PARAMETER To
        Recipient e-mail address list.
    .PARAMETER Subject
        Mail subject.
    .PARAMETER AttachmentPath
        Optional attachment path(s).
    .PARAMETER ForceMailTo
        Indicates the recipient was manually overridden.
    .PARAMETER LogLocation
        Optional log file location.
    .EXAMPLE
        Send-ReportIfChanged -SendLogs $true -ModCounter 3 -MailBody $mailBody -SmtpServer $smtp -From $from -To $to -Subject $subject -AttachmentPath $transcript
    .EXAMPLE
        $result = Send-ReportIfChanged -SendLogs $true -ModCounter 0 -MailBody $mailBody -SmtpServer $smtp -From $from -To $to -Subject $subject
        if (-not $result.Success) { throw $result.Message }
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/send-reportifchanged
    #>
    [CmdletBinding()]
    param(
        [bool]$SendLogs = $true,

        [Parameter(Mandatory = $true)]
        [int]$ModCounter,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$MailBody,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SmtpServer,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$From,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$To,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Subject,

        [string[]]$AttachmentPath,

        [bool]$ForceMailTo = $false,

        [string]$LogLocation
    )

    $result = [pscustomobject]@{
        Success     = $false
        Sent        = $false
        ModCounter  = $ModCounter
        Message     = $null
        FinalBody   = $null
    }

    try {
        $finalBody = $MailBody
        if ($finalBody -notmatch '</ul>\s*</body>\s*</html>\s*$') {
            $finalBody += '</ul></body></html>'
        }
        $result.FinalBody = $finalBody

        if (-not $SendLogs) {
            $result.Success = $true
            $result.Message = 'Report sending disabled by SendLogs flag.'
            if ($LogLocation) { Write-Log -LogLocation $LogLocation -Message $result.Message -Level INFO -WriteToFile }
            return $result
        }

        if ($ModCounter -le 0) {
            $result.Success = $true
            $result.Message = 'No changes made, nothing to do.'
            if ($LogLocation) { Write-Log -LogLocation $LogLocation -Message $result.Message -Level INFO -WriteToFile }
            return $result
        }

        if ($LogLocation) {
            Write-Log -LogLocation $LogLocation -Message "Sending e-mail report with $ModCounter modifications." -Level INFO -WriteToFile
            if ($ForceMailTo) {
                Write-Log -LogLocation $LogLocation -Message "Forcing e-mail to: $($To -join ', ')" -Level WARNING -WriteToFile
            }
        }

        Send-Mail -SMTPServer $SmtpServer -From $From -To $To -Subject $Subject -Body $finalBody -AttachmentPath $AttachmentPath

        $result.Success = $true
        $result.Sent = $true
        $result.Message = 'Report sent successfully.'
    } catch {
        $result.Message = $_.Exception.Message
        if ($LogLocation) { Write-Log -LogLocation $LogLocation -Message "Failed to send report: $($result.Message)" -Level ERROR -WriteToFile }
    }

    return $result
}

