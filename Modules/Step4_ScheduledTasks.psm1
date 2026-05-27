# =============================================================================
#  Step4_ScheduledTasks.psm1
#  Registrierte Windows-Tasks anstossen und Ausfuehrungsstatus zurueckmelden.
#  ASCII-only (PS 5.1, BOM-los sicher).
# =============================================================================

function Invoke-StartScheduledTask {
    [CmdletBinding()]
    param(
        [string]$TaskPath,
        [string]$TaskName,
        [string]$Label       = '',
        [string]$ResultFile  = '',
        $StatusTextBlock,
        $BrushSuccess,
        $BrushDanger
    )

    if ([string]::IsNullOrWhiteSpace($TaskPath) -or [string]::IsNullOrWhiteSpace($TaskName)) {
        Write-Log "Invoke-StartScheduledTask: TaskPath oder TaskName fehlt." -Level Error
        return
    }

    $displayLabel = if ([string]::IsNullOrWhiteSpace($Label)) { $TaskName } else { $Label }

    if ($StatusTextBlock) { $StatusTextBlock.Text = "Warte auf Bestaetigung..." }

    $params = @{
        TaskPath        = $TaskPath
        TaskName        = $TaskName
        Label           = $displayLabel
        ResultFile      = $ResultFile
        StatusTextBlock = $StatusTextBlock
    }

    $onComplete = {
        param($sync)

        if ($sync.Error) {
            Write-Log "Aufgabe '$displayLabel' Fehler: $($sync.Error.Exception.Message)" -Level Error
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fehler: $($sync.Error.Exception.Message)"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
            return
        }

        if ($sync.Cancelled) {
            Write-Log "Aufgabe '$displayLabel': Abgebrochen." -Level Info
            if ($StatusTextBlock) { $StatusTextBlock.Text = "Abgebrochen" }
            return
        }

        $code = $sync.Result
        if ($null -eq $code) {
            Write-Log "Aufgabe '$displayLabel': Timeout, kein Ergebniscode." -Level Warning
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Timeout - kein Ergebnis"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
            return
        }
        if ($code -eq 0) {
            Write-Log "Aufgabe '$displayLabel' erfolgreich abgeschlossen (Code 0)." -Level Success
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fertig (Code 0)"
                if ($BrushSuccess) { $StatusTextBlock.Foreground = $BrushSuccess }
            }
        } else {
            Write-Log "Aufgabe '$displayLabel' beendet mit Code $code." -Level Warning
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Beendet mit Code $code"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
        }
    }.GetNewClosure()

    Start-RunspaceJob -Parameters $params -OnComplete $onComplete -ScriptBlock {

        # --- Sicherheits-Rueckfrage ---
        $confirmed = Confirm-Dialog `
            -Message "Aufgabe '$Label' jetzt ausfuehren?" `
            -Title   "Aufgabenstart bestaetigen"

        if (-not $confirmed) {
            Write-Log "Start der Aufgabe '$Label' vom Benutzer abgebrochen." -Level Warning
            $Sync.Cancelled = $true
            return $null
        }

        # --- StatusTextBlock auf "Laeuft..." setzen (UI-Thread via Dispatcher) ---
        if ($StatusTextBlock) {
            $tb = $StatusTextBlock
            $cl = { $tb.Text = "Laeuft..." }.GetNewClosure()
            [void]$Sync.Dispatcher.Invoke([Action]$cl)
        }

        # --- ScheduledTasks-Modul pruefen / laden ---
        $hasCmdlet = $null -ne (Get-Command 'Start-ScheduledTask' -ErrorAction SilentlyContinue)
        if (-not $hasCmdlet) {
            Import-Module ScheduledTasks -ErrorAction SilentlyContinue
            $hasCmdlet = $null -ne (Get-Command 'Start-ScheduledTask' -ErrorAction SilentlyContinue)
        }
        $hasInfo = $null -ne (Get-Command 'Get-ScheduledTaskInfo' -ErrorAction SilentlyContinue)

        # --- Baseline LastRunTime sichern ---
        $baseline = $null
        if ($hasInfo) {
            try {
                $infoBase = Get-ScheduledTaskInfo -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
                $baseline = $infoBase.LastRunTime
                Write-Log "Baseline LastRunTime: $baseline" -Level Debug
            } catch {
                Write-Log "Baseline-Ermittlung fehlgeschlagen: $($_.Exception.Message)" -Level Warning
            }
        }

        # --- Task starten ---
        Write-Log "Starte Aufgabe '$Label' ($TaskPath$TaskName)" -Level Info
        try {
            if ($hasCmdlet) {
                Start-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
            } else {
                # Fallback: schtasks.exe /Run /TN - KEINE weiteren Argumente
                $tn   = "$TaskPath$TaskName"
                $proc = [System.Diagnostics.Process]::Start('schtasks.exe', "/Run /TN `"$tn`"")
                if ($proc) { [void]$proc.WaitForExit(10000) }
            }
        } catch {
            Write-Log "Fehler beim Starten der Aufgabe: $($_.Exception.Message)" -Level Error
            throw
        }
        Write-Log "Start-Befehl abgesetzt. Warte auf Abschluss..." -Level Info

        if (-not $hasInfo) {
            Write-Log "Get-ScheduledTaskInfo nicht verfuegbar, kein Polling moeglich." -Level Warning
            return $null
        }

        # --- Polling via Get-ScheduledTaskInfo (LastTaskResult / LastRunTime) ---
        # 0x41301 (267009) = Task laeuft gerade
        $timeoutSec = 300
        $elapsed    = 0
        $pollSec    = 3
        $result     = $null

        Start-Sleep -Seconds 2   # Anlaufzeit fuer den Task-Start

        while ($elapsed -lt $timeoutSec) {
            Start-Sleep -Seconds $pollSec
            $elapsed += $pollSec

            $info = $null
            try {
                $info = Get-ScheduledTaskInfo -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
            } catch {
                Write-Log "Polling-Fehler (${elapsed}s): $($_.Exception.Message)" -Level Warning
                break
            }

            $code        = $info.LastTaskResult
            $lastRunTime = $info.LastRunTime

            if ($code -eq 0x41301) {
                Write-Log "Aufgabe laeuft... (${elapsed}s)" -Level Debug
                continue
            }

            # LastRunTime > Baseline = neuer Lauf abgeschlossen
            $newRun = ($null -ne $lastRunTime) -and (
                ($null -eq $baseline) -or ($lastRunTime -gt $baseline)
            )

            if ($newRun) {
                $result = [int]$code
                Write-Log "Aufgabe abgeschlossen. LastTaskResult: $result  LastRunTime: $lastRunTime" -Level Info
                break
            }

            if ($null -eq $baseline) {
                $result = [int]$code
                Write-Log "Aufgabe abgeschlossen (kein Baseline). LastTaskResult: $result" -Level Info
                break
            }

            Write-Log "Warte auf neue LastRunTime (Code $code, ${elapsed}s)" -Level Debug
        }

        if ($elapsed -ge $timeoutSec -and $null -eq $result) {
            Write-Log "Timeout (${timeoutSec}s): Aufgabe '$Label' nicht abgeschlossen." -Level Warning
        }

        # --- Optionale ResultFile-Pruefung (nur Existenz + Aenderungszeit, kein Inhalt) ---
        if ($result -eq 0 -and $ResultFile -ne '') {
            if ([System.IO.File]::Exists($ResultFile)) {
                $lw = [System.IO.File]::GetLastWriteTime($ResultFile)
                Write-Log "ResultFile vorhanden, zuletzt geaendert: $lw" -Level Info
            } else {
                Write-Log "ResultFile nicht gefunden: $ResultFile" -Level Warning
            }
        }

        return $result
    } | Out-Null
}

# Wrapper: Task per Id aus Config.ScheduledTasks aufloesen
function Invoke-ScheduledTaskById {
    [CmdletBinding()]
    param(
        [string]$Id,
        $StatusTextBlock,
        $BrushSuccess,
        $BrushDanger
    )

    $task = $global:Config.ScheduledTasks |
        Where-Object { $_.Id -eq $Id } |
        Select-Object -First 1

    if (-not $task) {
        Write-Log "Kein ScheduledTask mit Id '$Id' in der Config." -Level Error
        return
    }

    $rf = if ($task.ResultFile) { [string]$task.ResultFile } else { '' }

    Invoke-StartScheduledTask `
        -TaskPath        ([string]$task.TaskPath) `
        -TaskName        ([string]$task.TaskName) `
        -Label           ([string]$task.Label) `
        -ResultFile      $rf `
        -StatusTextBlock $StatusTextBlock `
        -BrushSuccess    $BrushSuccess `
        -BrushDanger     $BrushDanger
}

Export-ModuleMember -Function *
