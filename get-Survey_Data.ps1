# === Parameter === 1120e446.ewe.de@emea.teams.ms


$Date = Get-Date -Format yyyy-MM
$DatelatMonth = (Get-Date).AddMonths(-1).ToString("yyyy-MM")
$Datemailmessages = if((Get-Date -Format ddd) -eq "Mo") { (Get-Date).AddDays(-3).ToString("M_yyyy") } Else {(Get-Date).AddDays(-1).ToString("M_yyyy")}
#Test
#$csvPath = "C:\Users\MRPLAGA\OneDrive - EWE Aktiengesellschaft\Dokumente\Genesys_OneDrive\Survey\Test\2025-08-13_Survey.csv"
#$csvPath = "C:\Users\MRPLAGA\OneDrive - EWE Aktiengesellschaft\Dokumente\Genesys_OneDrive\Survey\$DatelatMonth\Survey_Data-Last_Month.csv"
$csvPathFolder = "C:\Users\MRPLAGA\OneDrive - EWE Aktiengesellschaft\EN-NV-Kundenservice - EmailMessages_$Datemailmessages"
# 1) Erst mal alle Dateien im Ordner listen
# Finde die Datei, auch wenn die Zahl sich ändert
$file = Get-ChildItem -Path $csvPathFolder -Filter "*SurveyYesterday*.csv" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
$csvRows = Import-Csv -Path $file.FullName -Delimiter ";"

$subFolderName = if((Get-Date -Format ddd) -eq "Mo") { (Get-Date).AddDays(-3).ToString("yyyy-MM-dd") } Else {(Get-Date).AddDays(-1).ToString("yyyy-MM-dd")}
$newFolderPath = "C:\Users\MRPLAGA\OneDrive - EWE Aktiengesellschaft\Dokumente\Genesys_OneDrive\Survey\"+$Datemailmessages+"\"+$subFolderName
# Ordner erstellen (falls nicht vorhanden)
if (-not (Test-Path $newFolderPath)) {
    New-Item -Path $newFolderPath -ItemType Directory | Out-Null
}
#Test
#$outputCsvPath = "C:\Users\MRPLAGA\OneDrive - EWE Aktiengesellschaft\Dokumente\Genesys_OneDrive\Survey\Test\2025-08-13_Survey_Gesamt.csv"
$outputCsvPath = $newFolderPath+"\"+$subFolderName+"_Conversations_With_Survey_Gesamt.csv"
$apiBaseUrl = "https://api.mypurecloud.de/api/v2/quality/conversations"

# === Auth-Parameter ===
$clientId = "579762c3-7a27-41d9-b36b-a0fcafb04c82"
$clientSecret = "cB6MCT1V4swGkQjU1GIhBWncMfVNJlYq_x_jk_HRbKM"
$region = "mypurecloud.de"

# === OAuth holen ===
$tokenUrl = "https://login.$region/oauth/token"
$authBody = @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
}

$response = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $authBody -ContentType "application/x-www-form-urlencoded"

$accessToken = $response.access_token
#Write-Host "Token geholt: $accessToken"

# === Headers mit Token ===
$headers = @{
    "Authorization" = "Bearer $accessToken"
    "Content-Type"  = "application/json"
}

# === CSV einlesen ===
#$csvRows = Import-Csv -Path $csvPath -Delimiter ";"

# === Array für neue Zeilen ===
$newRows = @()

$requestCounter = 0
foreach ($row in $csvRows) {

    # === Eingangskanal bestimmen ===
    switch ($row.DNIS) {
        "tel:+4944135011621" { $eingangskanal = "MPlus" }
        "tel:+4944135011622" { $eingangskanal = "Teleteam" }
        "tel:+4944135011623" { $eingangskanal = "Test" }
        "tel:+4944135011698" { $eingangskanal = "EWE NETZ Hotline" }
        "tel:+4944135011696" { $eingangskanal = "EWE NETZ Hotline" }
        "tel:+4944135011695" { $eingangskanal = "BM" }
        "tel:+4944135011691" { $eingangskanal = "Marktpartner/Debitor" }
        default { $eingangskanal = "" }
    }

    # Spalte anhängen
    $row | Add-Member -NotePropertyName "Eingangskanal" -NotePropertyValue $eingangskanal
    # Vendor bestimmen
    $users = $row.'Users - Interacted'
    if ($users -match "Transcom") {
        $vendor = "Transcom"
    } elseif ($users -match "Teleperformance") {
        $vendor = "Teleperformance"
    } else {
        $vendor = ""
    }
    $row | Add-Member -NotePropertyName "Vendor" -NotePropertyValue $vendor
    # Prüfen ob Survey Status Finished ist
    if ($row."Survey Status" -eq "Finished") {

        $conversationId = $row."Conversation ID"

        # API URL zusammensetzen
        $url = "$apiBaseUrl/$conversationId/surveys"

        # API-Request
        try {
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method GET

            # === Antworten extrahieren ===
            $questionScores = $response.answers.questionGroupScores[0].questionScores

            $q1 = $questionScores | Where-Object { $_.questionId -eq "ed091d81-9271-442a-8fac-270dcfdeed34" }
            $q1Score = if ($q1) { $q1.score } else { "" }

            switch ($q1Score) {
                2 { $q1Text = "Ja, es wurde gelöst" }
                1 { $q1Text = "Nein, aber ich weiß wie es weitergeht" }
                0 { $q1Text = "Nein, es wurde nicht gelöst" }
                default { $q1Text = "" }
            }

            $q2 = $questionScores | Where-Object { $_.questionId -eq "36c5b145-3c60-48c4-916b-8a30bd75cf15" }
            $q2Score = if ($q2) { $q2.score } else { "" }

            $q3 = $questionScores | Where-Object { $_.questionId -eq "42fea1ca-6b3f-4604-8365-42510c5881a7" }
            $q3Text = if ($q3) { $q3.freeTextAnswer } else { "" }

            # Neue Spalten anhängen
            $row | Add-Member -NotePropertyName "Answer1_KonnteAnliegenGeloest" -NotePropertyValue $q1Score
            $row | Add-Member -NotePropertyName "Answer1_KonnteAnliegenGeloest_Text" -NotePropertyValue $q1Text
            $row | Add-Member -NotePropertyName "Answer2_Zufriedenheit" -NotePropertyValue $q2Score
            $row | Add-Member -NotePropertyName "Answer3_Hauptgrund" -NotePropertyValue $q3Text

            # === Requestzähler erhöhen ===
            $requestCounter++
            if ($requestCounter -ge 205) {
                Write-Host "Schwelle erreicht ($requestCounter Requests) – Warte 60 Sekunden..."
                Start-Sleep -Seconds 60
                $requestCounter = 0
            }

        } catch {
            Write-Warning "Fehler beim Abrufen von Survey für ConversationID: $conversationId"
        }

    } else {
        # Keine neuen Spalten, leer lassen
        $row | Add-Member -NotePropertyName "Answer1_KonnteAnliegenGeloest" -NotePropertyValue ""
        $row | Add-Member -NotePropertyName "Answer1_KonnteAnliegenGeloest_Text" -NotePropertyValue ""
        $row | Add-Member -NotePropertyName "Answer2_Zufriedenheit" -NotePropertyValue ""
        $row | Add-Member -NotePropertyName "Answer3_Hauptgrund" -NotePropertyValue ""
    }

    $newRows += $row
}

# === Neue CSV speichern === 
$csvText = $newRows | ConvertTo-Csv -NoTypeInformation -Delimiter ";" 
Set-Content -Path $outputCsvPath -Value $csvText -Encoding UTF8 -Force 
Write-Host "Neue CSV mit Survey-Antworten gespeichert: $outputCsvPath" 

# === Zusätzliche Tabellen erstellen === 
# Filter: Mplus 
$mplusRows = $newRows | Where-Object { $_.Eingangskanal -eq "MPlus" } 
$mplusPath = $outputCsvPath -replace "_Gesamt.csv$", "_MPlus.csv" 
$csvText = $mplusRows | ConvertTo-Csv -NoTypeInformation -Delimiter ";" 
Set-Content -Path $mplusPath -Value $csvText -Encoding UTF8 -Force 
Write-Host "MPlus CSV gespeichert: $mplusPath" 

# Filter: Teleteam 
$teleteamRows = $newRows | Where-Object { $_.Eingangskanal -eq "Teleteam" } 
$teleteamPath = $outputCsvPath -replace "_Gesamt.csv$", "_Teleteam.csv" 
$csvText = $teleteamRows | ConvertTo-Csv -NoTypeInformation -Delimiter ";" 
Set-Content -Path $teleteamPath -Value $csvText -Encoding UTF8 -Force 
Write-Host "Teleteam CSV gespeichert: $teleteamPath" 

# Filter: Transcom 
$transcomRows = $newRows | Where-Object { $_.Vendor -eq "Transcom" } 
$transcomPath = $outputCsvPath -replace "_Gesamt.csv$", "_Transcom.csv" 
$csvText = $transcomRows | ConvertTo-Csv -NoTypeInformation -Delimiter ";" 
Set-Content -Path $transcomPath -Value $csvText -Encoding UTF8 -Force 
Write-Host "Transcom CSV gespeichert: $transcomPath" 

# Filter: Teleperformance 
$teleperformanceRows = $newRows | Where-Object { $_.Vendor -eq "Teleperformance" } 
$teleperformancePath = $outputCsvPath -replace "_Gesamt.csv$", "_Teleperformance.csv" 
$csvText = $teleperformanceRows | ConvertTo-Csv -NoTypeInformation -Delimiter ";" 
Set-Content -Path $teleperformancePath -Value $csvText -Encoding UTF8 -Force 

Write-Host "Teleperformance CSV gespeichert: $teleperformancePath"

#Remove-Item "$($csvPathFolder)\$($file)" -Confirm:$false
#$number = $file.Name -replace '\D', ''
#Remove-Item "$($csvPathFolder)\Genesys Cloud Report from Plaga Markus_$($number).eml" -Confirm:$false

# ============================================================
# === Verarbeitung für SurveyLastWeek ========================
# ============================================================

# Datei suchen
$fileLastWeek = Get-ChildItem -Path $csvPathFolder -Filter "*SurveyLastWeek*.csv" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

if ($fileLastWeek) {
    $csvRowsLastWeek = Import-Csv -Path $fileLastWeek.FullName -Delimiter ";"

    # Kalenderwoche bestimmen (immer von letzter Woche)
    # Deutsch als Kultur einstellen
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo("de-DE")
    $calendar = $culture.Calendar

    # Gestern nehmen, um die vorherige Woche zu bekommen
    $yesterday = (Get-Date).AddDays(-7)

    $kwName  = "KS"+$calendar.GetWeekOfYear(
        $yesterday,
        [System.Globalization.CalendarWeekRule]::FirstFourDayWeek,
        [System.DayOfWeek]::Monday
    )



    # Ordnerpfad: ...\Survey\MM_yyyy\KWxx
    $newFolderPathLastWeek = "C:\Users\MRPLAGA\OneDrive - EWE Aktiengesellschaft\Dokumente\Genesys_OneDrive\Survey\"+$Datemailmessages+"\"+$kwName

    if (-not (Test-Path $newFolderPathLastWeek)) {
        New-Item -Path $newFolderPathLastWeek -ItemType Directory | Out-Null
    }

    # Dateiname: Conversations_LastWeek.csv
    $outputCsvPathLastWeek = $newFolderPathLastWeek+"\Conversations_With_Survey_LastWeek.csv"

    $newRowsLastWeek = @()
    $requestCounter = 0

    foreach ($row in $csvRowsLastWeek) {
        # Eingangskanal wie gehabt
        switch ($row.DNIS) {
            "tel:+4944135011621" { $eingangskanal = "MPlus" }
            "tel:+4944135011622" { $eingangskanal = "Teleteam" }
            "tel:+4944135011623" { $eingangskanal = "Test" }
            "tel:+4944135011698" { $eingangskanal = "EWE NETZ Hotline" }
            "tel:+4944135011696" { $eingangskanal = "EWE NETZ Hotline" }
            "tel:+4944135011695" { $eingangskanal = "BM" }
            "tel:+4944135011691" { $eingangskanal = "Marktpartner/Debitor" }
            default { $eingangskanal = "" }
        }
        $row | Add-Member -NotePropertyName "Eingangskanal" -NotePropertyValue $eingangskanal

        # Vendor bestimmen
        $users = $row.'Users - Interacted'
        if ($users -match "Transcom") {
            $vendor = "Transcom"
        } elseif ($users -match "Teleperformance") {
            $vendor = "Teleperformance"
        } else {
            $vendor = ""
        }
        $row | Add-Member -NotePropertyName "Vendor" -NotePropertyValue $vendor

        # Survey prüfen
        if ($row."Survey Status" -eq "Finished") {
            $conversationId = $row."Conversation ID"
            $url = "$apiBaseUrl/$conversationId/surveys"

            try {
                $response = Invoke-RestMethod -Uri $url -Headers $headers -Method GET
                $questionScores = $response.answers.questionGroupScores[0].questionScores

                $q1 = $questionScores | Where-Object { $_.questionId -eq "ed091d81-9271-442a-8fac-270dcfdeed34" }
                $q1Score = if ($q1) { $q1.score } else { "" }
                switch ($q1Score) {
                    2 { $q1Text = "Ja, es wurde gelöst" }
                    1 { $q1Text = "Nein, aber ich weiß wie es weitergeht" }
                    0 { $q1Text = "Nein, es wurde nicht gelöst" }
                    default { $q1Text = "" }
                }

                $q2 = $questionScores | Where-Object { $_.questionId -eq "36c5b145-3c60-48c4-916b-8a30bd75cf15" }
                $q2Score = if ($q2) { $q2.score } else { "" }

                $q3 = $questionScores | Where-Object { $_.questionId -eq "42fea1ca-6b3f-4604-8365-42510c5881a7" }
                $q3Text = if ($q3) { $q3.freeTextAnswer } else { "" }

                $row | Add-Member -NotePropertyName "Answer1_KonnteAnliegenGeloest" -NotePropertyValue $q1Score
                $row | Add-Member -NotePropertyName "Answer1_KonnteAnliegenGeloest_Text" -NotePropertyValue $q1Text
                $row | Add-Member -NotePropertyName "Answer2_Zufriedenheit" -NotePropertyValue $q2Score
                $row | Add-Member -NotePropertyName "Answer3_Hauptgrund" -NotePropertyValue $q3Text

                $requestCounter++
                if ($requestCounter -ge 205) {
                    Write-Host "Schwelle erreicht ($requestCounter Requests) – Warte 60 Sekunden..."
                    Start-Sleep -Seconds 60
                    $requestCounter = 0
                }
            } catch {
                Write-Warning "Fehler beim Abrufen von Survey für ConversationID: $conversationId"
            }
        } else {
            $row | Add-Member -NotePropertyName "Answer1_KonnteAnliegenGeloest" -NotePropertyValue ""
            $row | Add-Member -NotePropertyName "Answer1_KonnteAnliegenGeloest_Text" -NotePropertyValue ""
            $row | Add-Member -NotePropertyName "Answer2_Zufriedenheit" -NotePropertyValue ""
            $row | Add-Member -NotePropertyName "Answer3_Hauptgrund" -NotePropertyValue ""
        }

        $newRowsLastWeek += $row
    }

    # Speichern der neuen CSV
    $csvText = $newRowsLastWeek | ConvertTo-Csv -NoTypeInformation -Delimiter ";"
    Set-Content -Path $outputCsvPathLastWeek -Value $csvText -Encoding UTF8 -Force
    Write-Host "Neue LastWeek CSV mit Survey-Antworten gespeichert: $outputCsvPathLastWeek"

    # Filter-Exporte wie gehabt
    $mplusRows = $newRowsLastWeek | Where-Object { $_.Eingangskanal -eq "MPlus" }
    $mplusPath = $outputCsvPathLastWeek -replace "_LastWeek.csv$", "_MPlus_LastWeek.csv"
    $mplusRows | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Set-Content -Path $mplusPath -Encoding UTF8 -Force

    $teleteamRows = $newRowsLastWeek | Where-Object { $_.Eingangskanal -eq "Teleteam" }
    $teleteamPath = $outputCsvPathLastWeek -replace "_LastWeek.csv$", "_Teleteam_LastWeek.csv"
    $teleteamRows | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Set-Content -Path $teleteamPath -Encoding UTF8 -Force

    $transcomRows = $newRowsLastWeek | Where-Object { $_.Vendor -eq "Transcom" }
    $transcomPath = $outputCsvPathLastWeek -replace "_LastWeek.csv$", "_Transcom_LastWeek.csv"
    $transcomRows | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Set-Content -Path $transcomPath -Encoding UTF8 -Force

    $teleperformanceRows = $newRowsLastWeek | Where-Object { $_.Vendor -eq "Teleperformance" }
    $teleperformancePath = $outputCsvPathLastWeek -replace "_LastWeek.csv$", "_Teleperformance_LastWeek.csv"
    $teleperformanceRows | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Set-Content -Path $teleperformancePath -Encoding UTF8 -Force

    Write-Host "Alle LastWeek-Filter CSVs gespeichert."

    # Aufräumen
    #Remove-Item "$($csvPathFolder)\$($fileLastWeek)" -Confirm:$false
}



if($fileLastWeek)
{
    # SMTP-Parameter
    $IsBodyHtml = $true
    $SMTPServer = "smtp.ewenet.ewe.de"
    $SMTPPort   = 587  

    $From           = "markus.plaga@ewe-netz.de"
    $RecipientEmail = "markus.plaga@ewe-netz.de"
    $Subject        = "Survey Data $kwName"

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

    # === Dateien aus dem Unterordner anhängen ===
    $attachments = Get-ChildItem -Path $newFolderPathLastWeek -File

    foreach ($att in $attachments) {
        $MailMessage.Attachments.Add($att.FullName) | Out-Null
    }

    # Mail versenden
    $SMTPClient.Send($MailMessage)
}

<#
#--------TP--------
$From           = "markus.plaga@ewe-netz.de"
$RecipientEmail = "markus.plaga@ewe-netz.de"
$Subject        = "Survey Data $subFolderName"

# SMTP-Client-Objekt erstellen
$SMTPClientTP = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
$SMTPClientTP.EnableSsl = $true
$SMTPClientTP.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials  # NTLM/Kerberos

# E-Mail erstellen
$MailMessageTP = New-Object Net.Mail.MailMessage
$MailMessageTP.From = $From
$MailMessageTP.To.Add($RecipientEmail)
$MailMessageTP.Subject = $Subject
$MailMessageTP.IsBodyHtml = $true
$MailMessageTP.BodyEncoding = [System.Text.Encoding]::UTF8
$MailMessageTP.Body = @"
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

# === Dateien aus dem Unterordner anhängen ===
$attachments = Get-ChildItem -Path $teleperformancePath -File

foreach ($att in $attachments) {
    $MailMessageTP.Attachments.Add($att.FullName) | Out-Null
}

# Mail versenden
$SMTPClient.Send($MailMessageTP)

#>

#--------TC--------
$From           = "markus.plaga@ewe-netz.de"
$RecipientEmail = "markus.plaga@ewe-netz.de" #"Eric.Eggert.ext@ewe-netz.de"
#$CcRecipients   = @("markus.plaga@ewe-netz.de", "sascha.spothelfer@transcom.com", "kathleen.andres@transcom.com")   # mehrere CC-Empfänger
$Subject        = "Survey Data $subFolderName"

# SMTP-Client-Objekt erstellen
$SMTPClientTC = New-Object Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
$SMTPClientTC.EnableSsl = $true
$SMTPClientTC.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials  # NTLM/Kerberos

# E-Mail erstellen
$MailMessageTC = New-Object Net.Mail.MailMessage
$MailMessageTC.From = $From
$MailMessageTC.To.Add($RecipientEmail)


# CC-Adressen hinzufügen
foreach ($cc in $CcRecipients) {
    $MailMessageTC.CC.Add($cc)
}

$MailMessageTC.Subject = $Subject
$MailMessageTC.IsBodyHtml = $true
$MailMessageTC.BodyEncoding = [System.Text.Encoding]::UTF8
# Body setzen
$MailMessageTC.Body = @"
<html>
<head>
    <meta charset='UTF-8'>
    <style>
        body { font-family: Arial, sans-serif; }
        .container { padding: 10px; }
    </style>
</head>
<body>
    <div class="container">
        <p>Guten Morgen,</p>
        <p>im Anhang befindet sich nun der NKB-Report des gestrigen Tages.</p>
        <p>VG<br>Markus</p>
    </div>
</body>
</html>
"@

# === Dateien aus dem Unterordner anhängen ===
$attachments = Get-ChildItem -Path $transcomPath -File

foreach ($att in $attachments) {
    $MailMessageTC.Attachments.Add($att.FullName) | Out-Null
}

# Mail versenden
$SMTPClient.Send($MailMessageTC)