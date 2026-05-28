# =============================================================================
#  Step3_Pruflauf.psm1
#  Schritt 3 - Pruflauf in Prosos.
# =============================================================================

function Invoke-Step3Copy {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        Write-Log "Keine Bezeichnung zum Kopieren vorhanden." -Level Warning
        return
    }

    try {
        [System.Windows.Clipboard]::SetText($Text)
        Write-Log "Bezeichnung kopiert: $Text" -Level Success
    } catch {
        Write-Log "Clipboard-Fehler: $($_.Exception.Message)" -Level Error
    }
}

function Invoke-Step3ReadSumme {
    param(
        [string]$TaskFolder,
        $SummeTextBlock,
        $BrushSuccess,
        $BrushDanger
    )

    $xlsxPath = Join-Path $TaskFolder 'alleRechnungen.xlsx'

    if (-not [System.IO.File]::Exists($xlsxPath)) {
        Write-Log "alleRechnungen.xlsx nicht gefunden: $xlsxPath" -Level Warning
        if ($SummeTextBlock) {
            $SummeTextBlock.Text = "Datei nicht gefunden"
            if ($BrushDanger) { $SummeTextBlock.Foreground = $BrushDanger }
        }
        return
    }

    Write-Log "Lese Gesamtsumme aus J2: $xlsxPath" -Level Info
    if ($SummeTextBlock) { $SummeTextBlock.Text = "..." }

    $params = @{ XlsxPath = $xlsxPath }

    $onComplete = {
        param($sync)
        if ($sync.Error) {
            Write-Log "Gesamtsumme: Fehler beim Lesen der Excel-Datei." -Level Error
            if ($SummeTextBlock) {
                $SummeTextBlock.Text = "Lesefehler"
                if ($BrushDanger) { $SummeTextBlock.Foreground = $BrushDanger }
            }
            return
        }
        $val = $sync.Result
        if ($null -eq $val) {
            Write-Log "Gesamtsumme: J2 ist leer." -Level Warning
            if ($SummeTextBlock) {
                $SummeTextBlock.Text = "Kein Wert in J2"
                if ($BrushDanger) { $SummeTextBlock.Foreground = $BrushDanger }
            }
            return
        }
        $de  = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
        $fmt = ([double]$val).ToString('N2', $de)
        Write-Log "Gesamtsumme aller Rechnungen ist $fmt Euro" -Level Success
        if ($SummeTextBlock) {
            $SummeTextBlock.Text = "$fmt Euro"
            if ($BrushSuccess) { $SummeTextBlock.Foreground = $BrushSuccess }
        }
    }.GetNewClosure()

    Start-RunspaceJob -Parameters $params -OnComplete $onComplete -ScriptBlock {
        $excel    = $null
        $workbook = $null
        try {
            $excel               = New-Object -ComObject Excel.Application
            $excel.Visible       = $false
            $excel.DisplayAlerts = $false
            $workbook = $excel.Workbooks.Open($XlsxPath)
            $ws = $workbook.Worksheets.Item(1)
            $ws.Cells.Item(2, 10).Value2
        } finally {
            if ($workbook) {
                [void]$workbook.Close($false)
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
            }
            if ($excel) {
                [void]$excel.Quit()
                [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    } | Out-Null
}

function Invoke-Step3OpenProsos {
    $exe = 'C:\Program Files (x86)\PROSOZ Herten\OPEN PROSOZ\Anwendungen\OpenStarter.exe'

    if ([System.IO.File]::Exists($exe)) {
        [System.Diagnostics.Process]::Start($exe, '/app OpenClient.exe') | Out-Null
        Write-Log "Prosos gestartet." -Level Info
    } else {
        Write-Log "Prosos-Exe nicht gefunden: $exe" -Level Warning
    }
}

Export-ModuleMember -Function *
