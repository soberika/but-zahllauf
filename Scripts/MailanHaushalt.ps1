<#
    Robuste Outlook-Entwurf-Erstellung mit gewünschtem Absender
    - Öffnet Schreibfenster zur manuellen Prüfung vor dem Senden
    - Setzt Absender auf JC.Systembetreuung@kreis-meissen.de
#>

param(
    [string]$Bezeichnung = 'Zahllauf'
)

# Parameter / Konfiguration
$recipient   = "JC.Haushalt@kreis-meissen.de"
$subject     = "$Bezeichnung Prüflauf - steht zur weiteren Kontrolle bereit"
$desiredFrom = "jc.systembetreuung@kreis-meissen.de"

$inspector = $null
$mail      = $null
$account   = $null
$outlook   = $null

try {
    # Outlook holen (laufend bevorzugen → stabiler)
    Write-Host "Outlook COM-Objekt wird gesucht..." -ForegroundColor Cyan
    try {
        $outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
        Write-Host " → Laufendes Outlook gefunden" -ForegroundColor Green
    }
    catch {
        Write-Host " → Outlook nicht aktiv → starte neu..." -ForegroundColor Yellow
        $outlook = New-Object -ComObject Outlook.Application
        Write-Host " → Outlook neu gestartet" -ForegroundColor Green
    }

    # Neues MailItem
    $mail = $outlook.CreateItem(0)   # olMailItem = 0

    # Grundeinstellungen
    $mail.To       = $recipient
    $mail.Subject  = $subject
    $mail.HTMLBody = "<br><br>"      # Leerzeilen → Signatur kommt darunter

    # Account finden (case-insensitive)
    Write-Host "Suche nach Account: $desiredFrom" -ForegroundColor Cyan
    $account = $null
    foreach ($acc in $outlook.Session.Accounts) {
        if ($acc.SmtpAddress -and $acc.SmtpAddress.Trim().ToLower() -eq $desiredFrom) {
            $account = $acc
            break
        }
    }

    if ($account) {
        Write-Host " → Account gefunden: $($account.DisplayName) / $($account.SmtpAddress)" -ForegroundColor Green

        # Zuverlässiges Setzen via Reflection (bester Workaround)
        try {
            $bindingFlags = [System.Reflection.BindingFlags]::SetProperty
            $mail.GetType().InvokeMember("SendUsingAccount", $bindingFlags, $null, $mail, $account)
            Write-Host " → SendUsingAccount erfolgreich gesetzt (via Reflection)" -ForegroundColor Green
        }
        catch {
            Write-Warning "Reflection-Setzen fehlgeschlagen: $($_.Exception.Message)"
            try {
                $mail.SendUsingAccount = $account
                Write-Host " → Fallback: Direkte Zuweisung hat geklappt" -ForegroundColor Green
            }
            catch {
                Write-Warning "Auch direkte Zuweisung fehlgeschlagen → bleibt beim Standard-Absender"
            }
        }
    }
    else {
        Write-Warning "Kein passender Account gefunden → verwendet Standard-Absender!"
        Write-Host "  Verfügbare Accounts:" -ForegroundColor Yellow
        $outlook.Session.Accounts | ForEach-Object {
            Write-Host "    - $($_.DisplayName) | $($_.SmtpAddress)" -ForegroundColor DarkYellow
        }
    }

    # Signatur triggern (lädt die Standard-Signatur ins HTMLBody)
    $mail.HTMLBody = $mail.HTMLBody

    # ← Schreibfenster öffnen statt direkt senden
    # $false = nicht-modal → Outlook bleibt bedienbar, Skript kehrt sofort zurück
    $inspector = $mail.GetInspector
    $inspector.Display($false)

    Write-Host "Schreibfenster geöffnet!" -ForegroundColor Green
    Write-Host "Absender: $desiredFrom | Empfänger: $recipient" -ForegroundColor Cyan
    Write-Host "Bitte Mail prüfen und manuell senden." -ForegroundColor Yellow
}
catch {
    Write-Error "Kritischer Fehler beim Öffnen des Schreibfensters:`n$($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
}
finally {
    # WICHTIG: $inspector und $mail werden NICHT freigegeben –
    # sonst schließt sich das Outlook-Fenster sofort wieder!
    Write-Host "COM-Objekte werden freigegeben..." -ForegroundColor DarkGray
    Start-Sleep -Milliseconds 800
    @($account, $outlook) | ForEach-Object {
        if ($_ -ne $null) {
            try { [void][Runtime.Interopservices.Marshal]::ReleaseComObject($_) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Host "Cleanup abgeschlossen." -ForegroundColor DarkGray
}