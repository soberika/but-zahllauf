# =============================================================================
#  Step5_Abschluss.psm1
#  Schritt 5 - Abschluss.
# =============================================================================

function Invoke-Step5MailTemplate {
    param([string]$Bezeichnung)

    $scriptName = $global:Config.Paths.ScriptMail
    $scriptPath = Join-Path $global:AppRoot ("Scripts\" + $scriptName)

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        Write-Log "Mail-Skript nicht gefunden: $scriptPath" -Level Error
        return
    }

    Write-Log "Oeffne Mail-Vorlage fuer: $Bezeichnung" -Level Info

    $params = @{
        ScriptPath  = $scriptPath
        Bezeichnung = $Bezeichnung
    }

    Start-RunspaceJob -Parameters $params -OnComplete {
        param($sync)
        if ($sync.Error) {
            Write-Log "Mail-Vorlage: Fehler - $($sync.Error.Exception.Message)" -Level Error
        } else {
            Write-Log "Mail-Vorlage geoeffnet." -Level Success
        }
    } -ScriptBlock {
        & $ScriptPath -Bezeichnung $Bezeichnung
    } | Out-Null
}

function Invoke-Step5Finish {
    $simulate = [bool]$global:Config.Behavior.SimulateCleanup

    $ctx         = Get-ZahllaufContext
    $bezeichnung = $ctx.Bezeichnung -replace '[\\/:*?"<>|]', '_'
    $timestamp   = [DateTime]::Now.ToString('yyyyMMdd_HHmmss')
    $uniquePart  = [System.Guid]::NewGuid().ToString('N').Substring(0, 8)

    Write-Log "Abschluss gestartet$(if ($simulate) { ' [SIMULATION - kein echtes Loeschen]' })." -Level Info

    $params = @{
        Simulate    = $simulate
        MsgFolder   = $global:Config.Paths.MsgFolder
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
        $pfx          = if ($Simulate) { '[SIMULATION] ' } else { '' }
        $logDestFile  = [System.IO.Path]::Combine($LogDestDir, "${Bezeichnung}_${Timestamp}_${UniquePart}.log")

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

        # --- extracted_Attachements aufraumen ---
        $extractDir = [System.IO.Path]::Combine($MsgFolder, 'extracted_Attachements')
        if ([System.IO.Directory]::Exists($extractDir)) {
            $attachFiles = [System.IO.Directory]::GetFiles($extractDir, '*.*', [System.IO.SearchOption]::AllDirectories)
            foreach ($f in $attachFiles) {
                Write-Log "${pfx}Loesche Anhang: $([System.IO.Path]::GetFileName($f))" -Level $(if ($Simulate) { 'Warning' } else { 'Info' })
                if (-not $Simulate) { [System.IO.File]::Delete($f) }
            }
            Write-Log "${pfx}$($attachFiles.Count) Anhang-Datei(en) $(if ($Simulate) { 'wuerden geloescht' } else { 'geloescht' })." -Level Info
        } else {
            Write-Log "extracted_Attachements nicht gefunden unter: $extractDir" -Level Debug
        }

        # --- Logfile in Taskplaner-Ordner sichern ---
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

        Write-Log "$(if ($Simulate) { '[SIMULATION] ' })Abschluss abgeschlossen." -Level Success
    } | Out-Null
}

Export-ModuleMember -Function *
