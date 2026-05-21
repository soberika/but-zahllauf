# ============================================================
# viafintech Task-Ordner Skript
# Speichern als: UTF-8 mit BOM
# ============================================================

param(
    [string]$SourcePath = "\\VJC\GB1_Haushalt\OPEN_Listen_ab_2010\viafintech\_Temp\Task",
    [string]$TargetBase = "\\VJC\GB1_Haushalt\OPEN_Listen_ab_2010\viafintech\Zahllauf"
)

# Systemdatum
$date   = Get-Date
$yy     = $date.ToString("yy")
$mm     = $date.ToString("MM")
$dd     = $date.ToString("dd")
$yyyy   = $date.ToString("yyyy")

# Kalenderwoche (ISO 8601 – Montag als Wochenbeginn)
$cultureInfo = [System.Globalization.CultureInfo]::CurrentCulture
$kw = $cultureInfo.Calendar.GetWeekOfYear(
    $date,
    [System.Globalization.CalendarWeekRule]::FirstFourDayWeek,
    [System.DayOfWeek]::Monday
)

$kwFormatted       = $kw.ToString("D2")
$kwMinus1          = [Math]::Max(1, $kw - 1)
$kwMinus1Formatted = $kwMinus1.ToString("D2")

# ============================================================
# Pfade
# ============================================================
$sourcePath = $SourcePath
$targetBase = $TargetBase

$folderName    = "${yy}_${mm}_${dd} Fuer ${kwMinus1Formatted}.KW"
$newFolderPath = Join-Path $sourcePath $folderName

# ============================================================
# 1. Ordner erstellen
# ============================================================
if (-not (Test-Path $newFolderPath)) {
    New-Item -ItemType Directory -Path $newFolderPath | Out-Null
    Write-Host "Ordner erstellt: $newFolderPath" -ForegroundColor Green
} else {
    Write-Host "Ordner existiert bereits: $newFolderPath" -ForegroundColor Yellow
}

# ============================================================
# 1b. Zahlliste Header fixen – NUR die PDF mit diesem exakten Namen
#     Wird VOR dem Umbenennen ausgeführt → Originalname bleibt erhalten
# ============================================================
$zahllistePdf = Join-Path $sourcePath "ZahllisteHHSTGesamtBetr.pdf"
$fixScript    = Join-Path $PSScriptRoot "Fix-ZahllisteHeader.ps1"
if (Test-Path $zahllistePdf) {
    if (Test-Path $fixScript) {
        Write-Host "=== Zahlliste Header wird gefixt ===" -ForegroundColor Yellow
        & $fixScript -InputPdf $zahllistePdf
        Write-Host "Header-Fix abgeschlossen." -ForegroundColor Green
    } else {
        Write-Warning "Fix-ZahllisteHeader.ps1 nicht gefunden! (muss im selben Ordner wie dieses Skript liegen)"
    }
} else {
    Write-Host "Keine ZahllisteHHSTGesamtBetr.pdf gefunden – Fix wird übersprungen." -ForegroundColor Gray
}

# ============================================================
# 2. Dateien umbenennen
# ============================================================
$files = Get-ChildItem -Path $sourcePath -File

foreach ($file in $files) {
    $newFileName = "${yy}_${mm}_${dd}_${kwMinus1Formatted}_KW_$($file.Name)"
    Rename-Item -Path $file.FullName -NewName $newFileName -ErrorAction Stop
    Write-Host "Umbenannt: $($file.Name) -> $newFileName" -ForegroundColor Cyan
}

# ============================================================
# 2b. TXT → XLSX Konvertierung
# ============================================================
$excel = $null
$workbook = $null
$worksheet = $null

$latestTxtFile = Get-ChildItem -Path $sourcePath -Filter "*.txt" |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if ($null -eq $latestTxtFile) {
    Write-Host "Keine *.txt-Datei gefunden – Konvertierung übersprungen." -ForegroundColor Yellow
} else {
    $txtFilePath  = $latestTxtFile.FullName
    $xlsxFilePath = [System.IO.Path]::ChangeExtension($txtFilePath, ".xlsx")

    Write-Host "Konvertiere: $($latestTxtFile.Name) → XLSX ..." -ForegroundColor Yellow

    try {
        # TXT einlesen
        Write-Host "   [1/5] TXT lesen ..." -ForegroundColor Gray
        $content = Get-Content -Path $txtFilePath -Encoding Default

        # Excel starten
        Write-Host "   [2/5] Excel starten ..." -ForegroundColor Gray
        $excel               = New-Object -ComObject Excel.Application
        $excel.Visible       = $false
        $excel.DisplayAlerts = $false

        $workbook  = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)

        # Daten reinschreiben (alles zunächst als Text)
        Write-Host "   [3/5] Daten schreiben ($($content.Count) Zeilen) ..." -ForegroundColor Gray
        $row = 1
        foreach ($line in $content) {
            $columns = $line -split "`t"
            $col = 1
            foreach ($value in $columns) {
                $worksheet.Cells.Item($row, $col).Value2 = $value.Trim()
                $col++
            }
            $row++
        }

        $usedRange = $worksheet.UsedRange
        $lastRow   = $usedRange.Rows.Count

        # Formatierung & Zahlenkonvertierung
        Write-Host "   [4/5] Formatieren + Tabelle ..." -ForegroundColor Gray

        $euroFormat = '#.##0,00 €;[Rot]-#.##0,00 €'

        foreach ($colIndex in @(1,16,18)) {
            if ($colIndex -gt $worksheet.UsedRange.Columns.Count) { continue }

            $colRange = $worksheet.Range(
                $worksheet.Cells.Item(1, $colIndex),
                $worksheet.Cells.Item($lastRow, $colIndex)
            )

            # Zuerst Format setzen (NumberFormatLocal ist in DE-Umgebungen stabiler)
            $colRange.NumberFormatLocal = $euroFormat

            # TextToColumns – positionsbasiert, ohne leere Kommas
            $null = $colRange.TextToColumns(
                $worksheet.Cells.Item(1, $colIndex),  # Destination
                1,                                    # DataType = xlDelimited
                1,                                    # TextQualifier = xlDoubleQuote
                $false,                               # ConsecutiveDelimiter
                $true,                                # Tab
                $false,                               # Semicolon
                $false,                               # Comma
                $false,                               # Space
                $false,                               # Other
                "",                                   # OtherChar
                $null,                                # FieldInfo → null = Default (General → Number bei erkennbaren Zahlen)
                ",",                                  # DecimalSeparator
                "."                                   # ThousandsSeparator
            )
        }

        # Tabelle erstellen
        $null = $worksheet.ListObjects.Add(1, $usedRange, $null, 1)
        $worksheet.ListObjects.Item(1).Name       = "ZahllisteTabelle"
        $worksheet.ListObjects.Item(1).TableStyle = "TableStyleMedium2"

        $usedRange.Columns.AutoFit() | Out-Null
        $worksheet.Rows(1).Font.Bold = $true

        # Speichern
        Write-Host "   [5/5] Speichern ..." -ForegroundColor Gray
        $workbook.SaveAs($xlsxFilePath, 51)  # 51 = xlOpenXMLWorkbook (.xlsx)

        Write-Host "Fertig: $xlsxFilePath" -ForegroundColor Green
    }
    catch {
        Write-Host "FEHLER bei der Konvertierung:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        if ($_.Exception.InnerException) {
            Write-Host "Inner Exception: $($_.Exception.InnerException.Message)" -ForegroundColor DarkRed
        }
    }
    finally {
        # Excel sauber schließen
        if ($worksheet) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($worksheet) }
        if ($workbook)  {
            $workbook.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
        if ($excel)     {
            $excel.Quit()
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        }
        $worksheet = $workbook = $excel = $null
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }

    if (Test-Path $xlsxFilePath) {
        Remove-Item $txtFilePath -Force
        Write-Host "TXT gelöscht: $txtFilePath" -ForegroundColor Gray
    } else {
        Write-Host "Warnung: XLSX fehlt → TXT bleibt erhalten" -ForegroundColor Yellow
    }
}

# ============================================================
# 3. Dateien in neuen Ordner verschieben
# ============================================================
Get-ChildItem -Path $sourcePath -File | ForEach-Object {
    Move-Item -Path $_.FullName -Destination $newFolderPath -ErrorAction Stop
    Write-Host "Verschoben: $($_.Name)" -ForegroundColor Cyan
}

# ============================================================
# 4. Ordner ins Ziel verschieben
# ============================================================
$finalDestination = Join-Path $targetBase $folderName

try {
    if (Test-Path $finalDestination) {
        Write-Host "Zielordner existiert bereits – überspringe: $finalDestination" -ForegroundColor Yellow
    } else {
        Move-Item -Path $newFolderPath -Destination $finalDestination -ErrorAction Stop
        Write-Host "Ordner verschoben → $finalDestination" -ForegroundColor Green
    }
} catch {
    Write-Host "FEHLER beim finalen Verschieben: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nSkript beendet." -ForegroundColor White