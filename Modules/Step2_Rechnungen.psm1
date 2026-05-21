# =============================================================================
#  Step2_Rechnungen.psm1
#  Schritt 2 - ruft das Rechnungs-Skript im Runspace auf.
# =============================================================================

function Invoke-ExtractRechnungen {
    [CmdletBinding()]
    param(
        [switch]$NurExcel,
        [string]$CsvOrdner   = '',
        [string]$MsgFolder   = '',
        [string]$TaskFolder  = '',
        $StatusTextBlock,
        $BrushSuccess,
        $BrushDanger
    )

    $scriptName = $global:Config.Paths.ScriptRechnungen
    $scriptPath = Join-Path $global:AppRoot ("Scripts\" + $scriptName)

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        Write-Log "Rechnungs-Skript nicht gefunden: $scriptPath" -Level Error
        if ($StatusTextBlock) {
            $StatusTextBlock.Text = "Fehler: Skript nicht gefunden"
            if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
        }
        return
    }

    Write-Log "Starte Rechnungs-Skript: $scriptPath" -Level Info
    Write-Log ("Parameter: NurExcel={0} CsvOrdner='{1}' MsgFolder='{2}' TaskFolder='{3}'" -f `
        [bool]$NurExcel, $CsvOrdner, $MsgFolder, $TaskFolder) -Level Debug

    $params = @{
        ScriptPath = $scriptPath
        NurExcel   = [bool]$NurExcel
        CsvOrdner  = $CsvOrdner
        MsgFolder  = $MsgFolder
        TaskFolder = $TaskFolder
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

        $xlsx = Join-Path $TaskFolder 'alleRechnungen.xlsx'
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
        if ($NurExcel)                                       { $invokeArgs.NurExcel     = $true }
        if (-not [string]::IsNullOrWhiteSpace($CsvOrdner))   { $invokeArgs.CsvOrdner    = $CsvOrdner }
        if (-not [string]::IsNullOrWhiteSpace($MsgFolder))   { $invokeArgs.MsgFolder    = $MsgFolder }
        if (-not [string]::IsNullOrWhiteSpace($TaskFolder))  { $invokeArgs.TargetFolder = $TaskFolder }

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

    $effectiveNurExcel = [bool]($NurExcel -or $NurExcelCsv)

    Invoke-ExtractRechnungen `
        -NurExcel:$effectiveNurExcel `
        -CsvOrdner  $CsvOrdner `
        -MsgFolder  $global:Config.Paths.MsgFolder `
        -TaskFolder $global:Config.Paths.TaskFolder `
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
