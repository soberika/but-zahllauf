# =============================================================================
#  Step2_Rechnungen.psm1
#  Schritt 2 - ruft das Rechnungs-Skript (alleRechnungen / RechnungenausMailinExcel)
#  im Runspace auf und meldet das Ergebnis an die GUI zurueck.
# =============================================================================

function Invoke-ExtractRechnungen {
    [CmdletBinding()]
    param(
        # Wenn gesetzt, wird das Skript mit -NurExcel gestartet (CSVs muessen
        # dann bereits vorhanden sein).
        [switch]$NurExcel,

        # Optionaler Pfad zu einem alternativen CSV-Ordner.
        [string]$CsvOrdner = '',

        # UI-Status-TextBlock fuer Erfolgs-/Fehlermeldung (optional).
        $StatusTextBlock,

        # Brushes aus den Window-Resources (Erfolg gruen, Fehler rot).
        $BrushSuccess,
        $BrushDanger
    )

    $scriptName = $Script:Config.Paths.ScriptRechnungen
    $scriptPath = Join-Path $Script:AppRoot ("Scripts\" + $scriptName)

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        Write-Log "Rechnungs-Skript nicht gefunden: $scriptPath" -Level Error
        if ($StatusTextBlock) {
            $StatusTextBlock.Text = "Fehler: Skript nicht gefunden"
            if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
        }
        return
    }

    Write-Log "Starte Rechnungs-Skript: $scriptPath" -Level Info
    Write-Log "Parameter: NurExcel=$([bool]$NurExcel) CsvOrdner='$CsvOrdner'" -Level Debug

    $params = @{
        ScriptPath = $scriptPath
        NurExcel   = [bool]$NurExcel
        CsvOrdner  = $CsvOrdner
    }

    $onComplete = {
        param($sync)
        if ($sync.Error) {
            Write-Log "Schritt 2 abgebrochen." -Level Error
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fehler: $($sync.Error.Exception.Message)"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
            return
        }

        # Standard-Ablage des Skripts: Scripts\extracted_attachments\alleRechnungen.xlsx
        $defaultCsv = if ([string]::IsNullOrWhiteSpace($CsvOrdner)) {
            Join-Path $Script:AppRoot 'Scripts\extracted_attachments'
        } else { $CsvOrdner }
        $xlsx = Join-Path $defaultCsv 'alleRechnungen.xlsx'

        if (Test-Path -LiteralPath $xlsx) {
            Write-Log "Schritt 2 fertig. Excel: $xlsx" -Level Success
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fertig: $xlsx"
                if ($BrushSuccess) { $StatusTextBlock.Foreground = $BrushSuccess }
            }
        } else {
            Write-Log "Schritt 2 beendet, aber alleRechnungen.xlsx nicht gefunden unter: $xlsx" -Level Warning
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Beendet, XLSX nicht gefunden"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
        }
    }.GetNewClosure()

    Start-RunspaceJob -Parameters $params -OnComplete $onComplete -ScriptBlock {
        Write-Log "Extrahiere Anhaenge / erstelle Excel..." -Level Info

        $invokeArgs = @{}
        if ($NurExcel) { $invokeArgs.NurExcel = $true }
        if (-not [string]::IsNullOrWhiteSpace($CsvOrdner)) {
            $invokeArgs.CsvOrdner = $CsvOrdner
        }

        & $ScriptPath @invokeArgs
        Write-Log "Rechnungs-Skript durchgelaufen." -Level Info
    } | Out-Null
}

function Invoke-Step2Run {
    [CmdletBinding()]
    param(
        [switch]$NurExcel,
        [switch]$NurExcelCsv,
        [string]$CsvOrdner = '',
        $StatusTextBlock,
        $BrushSuccess,
        $BrushDanger
    )

    # Die UI hat zwei Checkboxen ("Nur Excel" und "Nur Excel + eigener CSV-Pfad").
    # Beide fuehren intern auf -NurExcel des Skripts.
    $effectiveNurExcel = [bool]($NurExcel -or $NurExcelCsv)

    Invoke-ExtractRechnungen `
        -NurExcel:$effectiveNurExcel `
        -CsvOrdner $CsvOrdner `
        -StatusTextBlock $StatusTextBlock `
        -BrushSuccess $BrushSuccess `
        -BrushDanger $BrushDanger
}

function Invoke-Step2OpenPath {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Start-Process explorer.exe -ArgumentList $Path
        Write-Log "Ergebnis-Ordner geoeffnet: $Path" -Level Info
    } else {
        Write-Log "Pfad nicht erreichbar: $Path" -Level Warning
    }
}

Export-ModuleMember -Function *
