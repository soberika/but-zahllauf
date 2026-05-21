# =============================================================================
#  Step4_Auszahlung.psm1
#  Schritt 4 - ruft viafintech_Task_Skript.ps1 im Runspace auf.
# =============================================================================

function Invoke-StartTaskScript {
    [CmdletBinding()]
    param(
        [string]$TaskFolder     = '',
        [string]$ZahllaufFolder = '',
        $StatusTextBlock,
        $BrushSuccess,
        $BrushDanger
    )

    $scriptName = $global:Config.Paths.ScriptTask
    $scriptPath = Join-Path $global:AppRoot ("Scripts\" + $scriptName)

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        Write-Log "Task-Skript nicht gefunden: $scriptPath" -Level Error
        if ($StatusTextBlock) {
            $StatusTextBlock.Text = "Fehler: Skript nicht gefunden"
            if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
        }
        return
    }

    $today          = Get-Date
    $kw             = Get-IsoCalendarWeek -Date $today
    $kwMinus1       = [Math]::Max(1, $kw - 1)
    $expectedFolder = "{0} Fuer {1:D2}.KW" -f $today.ToString('yy_MM_dd'), $kwMinus1

    Write-Log "Erwarteter Task-Ordner: $expectedFolder" -Level Info
    Write-Log "Starte Task-Skript: $scriptPath" -Level Info
    Write-Log "TaskFolder='$TaskFolder' ZahllaufFolder='$ZahllaufFolder'" -Level Debug

    if ($StatusTextBlock) { $StatusTextBlock.Text = "Laeuft..." }

    $params = @{
        ScriptPath     = $scriptPath
        ExpectedFolder = $expectedFolder
        TaskFolder     = $TaskFolder
        ZahllaufFolder = $ZahllaufFolder
    }

    $onComplete = {
        param($sync)
        if ($sync.Error) {
            Write-Log "Schritt 4 abgebrochen." -Level Error
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fehler: $($sync.Error.Exception.Message)"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
            return
        }

        $finalPath = Join-Path $ZahllaufFolder $expectedFolder
        if (Test-Path -LiteralPath $finalPath) {
            Write-Log "Schritt 4 fertig. Zahllauf-Ordner: $finalPath" -Level Success
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fertig: $finalPath"
                if ($BrushSuccess) { $StatusTextBlock.Foreground = $BrushSuccess }
            }
        } else {
            Write-Log "Schritt 4 beendet, aber Zielordner fehlt: $finalPath" -Level Warning
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Beendet, Ordner fehlt"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
        }
    }.GetNewClosure()

    Start-RunspaceJob -Parameters $params -OnComplete $onComplete -ScriptBlock {
        Write-Log "Erzeuge Ordner / benenne Dateien um / konvertiere TXT->XLSX..." -Level Info

        $invokeArgs = @{}
        if (-not [string]::IsNullOrWhiteSpace($TaskFolder))     { $invokeArgs.SourcePath = $TaskFolder }
        if (-not [string]::IsNullOrWhiteSpace($ZahllaufFolder)) { $invokeArgs.TargetBase = $ZahllaufFolder }

        & $ScriptPath @invokeArgs
        Write-Log "Task-Skript durchgelaufen." -Level Info
    } | Out-Null
}

function Invoke-Step4Run {
    [CmdletBinding()]
    param(
        $StatusTextBlock,
        $BrushSuccess,
        $BrushDanger
    )
    Invoke-StartTaskScript `
        -TaskFolder      $global:Config.Paths.TaskFolder `
        -ZahllaufFolder  $global:Config.Paths.ZahllaufFolder `
        -StatusTextBlock $StatusTextBlock `
        -BrushSuccess $BrushSuccess `
        -BrushDanger $BrushDanger
}

function Invoke-Step4OpenTask {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Start-Process explorer.exe -ArgumentList $Path
        Write-Log "Task-Ordner geoeffnet: $Path" -Level Info
    } else {
        Write-Log "Task-Ordner nicht erreichbar: $Path" -Level Warning
    }
}

function Invoke-Step4CheckCorrection {
    Write-Log "Letzte Korrektur pruefen (Platzhalter)." -Level Info
}

Export-ModuleMember -Function *
