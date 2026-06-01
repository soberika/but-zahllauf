# =============================================================================
#  Step5_Abschluss.psm1
#  Schritt 5 - Abschluss.
# =============================================================================

function Invoke-Step5MailTemplate {
    param(
        [string]$Bezeichnung,
        [string]$Summe        = '',
        [string]$ZielpfadLink = ''
    )

    $scriptName = $global:Config.Paths.ScriptMail
    $scriptPath = Join-Path $global:AppRoot ("Scripts\" + $scriptName)

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        Write-Log "Mail-Skript nicht gefunden: $scriptPath" -Level Error
        return
    }

    Write-Log "Oeffne Mail-Vorlage fuer: $Bezeichnung" -Level Info

    $params = @{
        ScriptPath    = $scriptPath
        Bezeichnung   = $Bezeichnung
        Summe         = $Summe
        ZielpfadLink  = $ZielpfadLink
    }

    Start-RunspaceJob -Parameters $params -OnComplete {
        param($sync)
        if ($sync.Error) {
            Write-Log "Mail-Vorlage: Fehler - $($sync.Error.Exception.Message)" -Level Error
        } else {
            Write-Log "Mail-Vorlage geoeffnet." -Level Success
        }
    } -ScriptBlock {
        & $ScriptPath -Bezeichnung $Bezeichnung -Summe $Summe -ZielpfadLink $ZielpfadLink
    } | Out-Null
}

# Loescht die temporaeren Dateien (.msg + extrahierte Anhaenge). Respektiert
# Behavior.SimulateCleanup (true = nur protokollieren, nicht loeschen).
function Invoke-Step5DeleteTemp {
    $simulate = [bool]$global:Config.Behavior.SimulateCleanup

    $extractedFolder = [string]$global:Config.Paths.ExtractedFolder
    if ([string]::IsNullOrWhiteSpace($extractedFolder)) {
        $extractedFolder = Join-Path $global:Config.Paths.TempRoot 'extracted_attachments'
    }

    Write-Log "Loeschen der temporaeren Dateien gestartet$(if ($simulate) { ' [SIMULATION - kein echtes Loeschen]' })." -Level Info

    $params = @{
        Simulate        = $simulate
        MsgFolder       = $global:Config.Paths.MsgFolder
        ExtractedFolder = $extractedFolder
    }

    Start-RunspaceJob -Parameters $params -OnComplete {
        param($sync)
        if ($sync.Error) {
            Write-Log "Loeschen fehlgeschlagen: $($sync.Error.Exception.Message)" -Level Error
        }
    } -ScriptBlock {
        $pfx = if ($Simulate) { '[SIMULATION] ' } else { '' }

        # --- .msg-Dateien loeschen ---
        if ([System.IO.Directory]::Exists($MsgFolder)) {
            $msgFiles = [System.IO.Directory]::GetFiles($MsgFolder, '*.msg')
            foreach ($f in $msgFiles) {
                Write-Log "${pfx}Loesche MSG: $([System.IO.Path]::GetFileName($f))" -Level $(if ($Simulate) { 'Warning' } else { 'Info' })
                if (-not $Simulate) { [System.IO.File]::Delete($f) }
            }
            Write-Log "${pfx}$($msgFiles.Count) MSG-Datei(en) $(if ($Simulate) { 'wuerden geloescht' } else { 'geloescht' })." -Level Info
        } else {
            Write-Log "MsgFolder nicht erreichbar: $MsgFolder" -Level Warning
        }

        # --- extrahierte Anhaenge (.pdf/.csv) aufraumen ---
        if ([System.IO.Directory]::Exists($ExtractedFolder)) {
            $attachFiles = [System.IO.Directory]::GetFiles($ExtractedFolder, '*.*', [System.IO.SearchOption]::AllDirectories)
            foreach ($f in $attachFiles) {
                Write-Log "${pfx}Loesche Anhang: $([System.IO.Path]::GetFileName($f))" -Level $(if ($Simulate) { 'Warning' } else { 'Info' })
                if (-not $Simulate) { [System.IO.File]::Delete($f) }
            }
            Write-Log "${pfx}$($attachFiles.Count) Anhang-Datei(en) $(if ($Simulate) { 'wuerden geloescht' } else { 'geloescht' })." -Level Info
        } else {
            Write-Log "Ordner mit extrahierten Anhaengen nicht gefunden: $ExtractedFolder" -Level Debug
        }

        Write-Log "$(if ($Simulate) { '[SIMULATION] ' })Temporaere Dateien aufgeraeumt." -Level Success
    } | Out-Null
}

# Schliesst den Zahllauf ab: sichert das Tageslog in den Taskplaner-Ordner und
# protokolliert, dass die Uebergabe veranlasst ist (Rueckmeldung Haushalt abwarten).
function Invoke-Step5Finish {
    param([string]$Bezeichnung = 'Zahllauf')

    $bezeichnung = $Bezeichnung -replace '[\\/:*?"<>|]', '_'
    $timestamp   = [DateTime]::Now.ToString('yyyyMMdd_HHmmss')
    $uniquePart  = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)

    Write-Log "Zahllauf abgeschlossen - sichere Log und veranlasse Uebergabe." -Level Info

    $params = @{
        Bezeichnung = $bezeichnung
        Timestamp   = $timestamp
        UniquePart  = $uniquePart
        LogDestDir  = '\\VJC\GB1_Systembetreuung\Taskplaner'
    }

    Start-RunspaceJob -Parameters $params -OnComplete {
        param($sync)
        if ($sync.Error) {
            Write-Log "Abschluss-Fehler: $($sync.Error.Exception.Message)" -Level Error
        }
    } -ScriptBlock {
        $logDestFile = [System.IO.Path]::Combine($LogDestDir, "${Bezeichnung}_${Timestamp}_${UniquePart}.log")

        Write-Log "Sichere Log nach: $logDestFile" -Level Info
        if ([System.IO.Directory]::Exists($LogDestDir)) {
            try {
                [System.IO.File]::Copy($Sync.LogFile, $logDestFile, $false)
                Write-Log "Log gesichert: $([System.IO.Path]::GetFileName($logDestFile))" -Level Success
            } catch {
                Write-Log "Log-Sicherung fehlgeschlagen: $($_.Exception.Message)" -Level Error
            }
        } else {
            Write-Log "Taskplaner-Ordner nicht erreichbar: $LogDestDir" -Level Warning
        }

        Write-Log "Zahllauf abgeschlossen. Uebergabe veranlasst, Log gesichert - Rueckmeldung vom Haushalt abwarten." -Level Success
    } | Out-Null
}

Export-ModuleMember -Function *
