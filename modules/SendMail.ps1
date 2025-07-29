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

    Modification History:
    - 2025-07-24: Fixed "$config" variable in the example and description of the Send-Mail function (typos, I don't use a config file).
                  Set mandatory parameters for the Send-Mail function.
    - 2025-05-23: Added CC parameter to the Send-Mail function.
    - 2025-04-03: Added Parameter informations to this help.
    - 2025-03-27: Initial version (isolation of e-mail functions from main script).
#>

# Mail Function
function Send-Mail {
    param (
        [Parameter(Mandatory = $true)][string] $SMTPServer,
        [int] $SMTPPort = 25,
        [Parameter(Mandatory = $true)][string] $From,
        [Parameter(Mandatory = $true)][string] $To,
        [string] $Cc = "",
        [string] $Bcc = "",
        [Parameter(Mandatory = $true)][string] $Subject,
        [Parameter(Mandatory = $true)][string] $Body,
        [string] $AttachmentPath
    )

    try {
        $mailMessage = New-Object System.Net.Mail.MailMessage($From, $To, $Subject, $Body)
        $mailMessage.IsBodyHtml = $true
        if ($AttachmentPath -ne "") { 
            $attachment = New-Object System.Net.Mail.Attachment($AttachmentPath)
            $mailMessage.Attachments.Add($attachment)
        }
        if ($Cc -ne "") { $mailMessage.CC.Add($Cc) }
        if ($Bcc -ne "") { $mailMessage.Bcc.Add($Bcc) }
        $smtpClient = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
        $smtpClient.Send($mailMessage)
    } catch {
        Write-Error "Failed to send e-mail: $_"
    }
}