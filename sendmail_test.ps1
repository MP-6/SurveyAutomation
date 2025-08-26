    # SMTP-Parameter
    $IsBodyHtml = $true
    $SMTPServer = "smtp.ewenet.ewe.de"
    $SMTPPort   = 587  

    $From           = "markus.plaga@ewe-netz.de"
    $RecipientEmail = "markus.plaga@ewe-netz.de"
    $Subject        = "Survey Data"

    # SMTP-Client-Objekt erstellen
    $SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials  # NTLM/Kerberos

    # E-Mail erstellen
    $MailMessage = New-Object Net.Mail.MailMessage
    $MailMessage.From = $From
    $MailMessage.To.Add($RecipientEmail)
    $MailMessage.Subject = $Subject
    $MailMessage.IsBodyHtml = $true
    $MailMessage.BodyEncoding = [System.Text.Encoding]::UTF8
    $MailMessage.Body = @"
    <html>
    <head>
        <meta charset='UTF-8'>
        <style>
            body { font-family: Arial, sans-serif; }
            .container { padding: 10px; }
            .highlight { font-weight: bold; color: #0078D7; }
            .footer { margin-top: 20px; font-style: italic; }
        </style>
    </head>
    <body>
    </body>
    </html>
"@

    # Mail versenden
    $SMTPClient.Send($MailMessage)