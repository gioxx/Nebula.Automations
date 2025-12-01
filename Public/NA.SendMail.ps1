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
    Send-Mail -From "lM2tH@example.com" -To "lM2tH@example.com" -Subject "Test e-mail" -Body "This is a test e-mail." -AttachmentPath "C:\test.txt"
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
