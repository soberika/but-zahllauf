# ============================================================
# alleRechnungen.ps1
#
# Was macht dieses Skript?
#   1. Liest alle .msg-E-Mails aus einem Netzwerkordner
#   2. Extrahiert die CSV-Anhaenge in einen lokalen Ordner
#   3. Kombiniert alle CSVs zu einer formatierten Excel-Datei
#   4. Bietet an, die fertige Excel ins Netzwerk zu verschieben
#
# Voraussetzungen:
#   - PowerShell 5.1 oder neuer
#   - Zwei Module: "ReadMsgFile" und "ImportExcel"
#     (Das Skript fragt automatisch nach ob es diese installieren soll)
#
# Aufruf-Varianten:
#   .\alleRechnungen.ps1
#       Kompletter Durchlauf: MSG-Extraktion + Excel-Erstellung
#
#   .\alleRechnungen.ps1 -NurExcel
#       Nur Excel erstellen (wenn CSVs bereits vorhanden sind)
#
#   .\alleRechnungen.ps1 -NurExcel -CsvOrdner "C:\MeinOrdner"
#       Nur Excel erstellen, CSVs aus einem eigenen Ordner lesen
# ============================================================

param(
    [switch]$NurExcel,
    [string]$CsvOrdner    = "",
    [string]$MsgFolder    = "\\vjc\GB1_Haushalt\OPEN_Listen_ab_2010\viafintech\_Temp",
    [string]$TargetFolder = "\\vjc\GB1_Haushalt\OPEN_Listen_ab_2010\viafintech\_Temp"
)

# $PSScriptRoot = Ordner in dem das Skript selbst liegt.
# Alle relativen Pfade werden davon abgeleitet, egal von wo
# das Skript gestartet wird (Taskplaner, Explorer, ISE, etc.)
$scriptDir = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($scriptDir)) {
    # Fallback falls das Skript direkt in der Konsole eingefuegt wird
    $scriptDir = (Get-Location).Path
}

# Standard-Ausgabeordner liegt neben dem Skript
if ([string]::IsNullOrWhiteSpace($CsvOrdner)) {
    $CsvOrdner = Join-Path $scriptDir "extracted_attachments"
}

# ────────────────────────────────────────────────────────────────
# HILFSFUNKTION: Modul pruefen und ggf. installieren
# ────────────────────────────────────────────────────────────────

function Ensure-Module {
    param([string]$ModulName)

    if (Get-Module -ListAvailable -Name $ModulName) {
        Import-Module $ModulName
        return
    }

    Write-Host ""
    Write-Host "Das Modul '$ModulName' ist nicht installiert." -ForegroundColor Yellow
    Write-Host "Es wird benoetigt damit das Skript funktioniert."
    Write-Host ""
    $antwort = Read-Host "Jetzt installieren? (Enter = Ja, 'nein' = Abbrechen)"

    if ($antwort -eq "nein" -or $antwort -eq "n") {
        Write-Host ""
        Write-Host "Installation abgebrochen. Manuell installieren mit:" -ForegroundColor Red
        Write-Host "   Install-Module -Name $ModulName -Scope CurrentUser" -ForegroundColor Cyan
        Write-Host ""
        exit 1
    }

    Write-Host "Installiere '$ModulName'..." -ForegroundColor Cyan
    try {
        Install-Module -Name $ModulName -Scope CurrentUser -Force -ErrorAction Stop
        Import-Module $ModulName
        Write-Host "'$ModulName' erfolgreich installiert." -ForegroundColor Green
    } catch {
        Write-Host "Installation fehlgeschlagen: $_" -ForegroundColor Red
        Write-Host "Versuche es manuell: Install-Module -Name $ModulName -Scope CurrentUser" -ForegroundColor Cyan
        exit 1
    }
}

# ────────────────────────────────────────────────────────────────
# TEIL 1: .msg-Dateien verarbeiten und Anhaenge extrahieren
# Wird uebersprungen wenn -NurExcel angegeben wurde
# ────────────────────────────────────────────────────────────────

if (-not $NurExcel) {

    Ensure-Module "ReadMsgFile"

    # Quellordner mit den .msg-E-Mails
    $msgFolder = $MsgFolder

    # Ausgabeordner liegt neben dem Skript (nicht im Arbeitsverzeichnis)
    New-Item -ItemType Directory -Path $CsvOrdner -Force | Out-Null
    Write-Host "Ausgabeordner: $CsvOrdner"

    $msgFiles = Get-ChildItem -Path $msgFolder -Filter "*.msg"

    if ($msgFiles.Count -eq 0) {
        Write-Host ""
        Write-Host "Keine .msg-Dateien gefunden in: $msgFolder" -ForegroundColor Yellow
        Write-Host "Bitte pruefen ob der Ordner erreichbar ist und Dateien enthaelt."
        exit 1
    }

    Write-Host ""
    Write-Host "$($msgFiles.Count) .msg-Datei(en) gefunden. Starte Extraktion..." -ForegroundColor Cyan
    Write-Host ""

    foreach ($file in $msgFiles) {
        $msgPath = $file.FullName
        $msgName = $file.BaseName

        try {
            $msg = Read-MsgFile -File $msgPath

            Write-Host "Verarbeite: ${msgPath}"
            Write-Host "  Betreff : $($msg.Subject)"
            Write-Host "  Von     : $($msg.From)"
            Write-Host "  Datum   : $($msg.Date)"

            # Anhaenge extrahieren - nur echte Dateiattachments, keine eingebetteten Nachrichten
            $attachments = $msg.Attachments | Where-Object { $_ -isnot [MsgReader.Outlook.Storage+Message] }

            if (-not $attachments -or @($attachments).Count -eq 0) {
                Write-Host "  (Keine Dateiattachments in dieser Mail)" -ForegroundColor Yellow
            } else {
                Get-MsgAttachment -File $msgPath -Path $CsvOrdner | ForEach-Object {
                    $tempPath = $_
                    if ($tempPath) {
                        $originalFilename = (Split-Path $tempPath -Leaf)
                        $newFilename = "${msgName}_${originalFilename}"
                        $newPath = Join-Path $CsvOrdner $newFilename
                        Move-Item -Path $tempPath -Destination $newPath -Force
                        Write-Host "  Anhang gespeichert: $newPath"
                    }
                }
            }

        } catch {
            Write-Host "  Fehler bei ${msgPath}: $_" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-Host "Extraktion abgeschlossen. Starte CSV-Verarbeitung..." -ForegroundColor Green
    Write-Host ""

}

# ────────────────────────────────────────────────────────────────
# TEIL 2: CSV-Dateien einlesen und zu Excel zusammenfuehren
# ────────────────────────────────────────────────────────────────

Ensure-Module "ImportExcel"

$excelOutputFile = Join-Path $CsvOrdner "alleRechnungen.xlsx"

$fixedHeaders = @(
    "type", "id", "reference_id", "currencycode",
    "net_amount", "vat_amount", "gross_amount",
    "settlement_reference_time", "settlement_event"
)

$csvFiles = Get-ChildItem -Path $CsvOrdner -Filter "*.csv"

if ($csvFiles.Count -eq 0) {
    Write-Host "Keine CSV-Dateien gefunden in: $CsvOrdner" -ForegroundColor Yellow
    Write-Host "Tipp: Pruefen ob die MSG-Extraktion Anhaenge gefunden hat."
    exit 1
}

Write-Host "$($csvFiles.Count) CSV-Datei(en) gefunden. Lese Daten..." -ForegroundColor Cyan
Write-Host ""

$allData = [System.Collections.Generic.List[object]]::new()

foreach ($file in $csvFiles) {
    $filePath = $file.FullName
    try {
        $data = Get-Content -Path $filePath |
                Select-Object -Skip 1 |
                ConvertFrom-Csv -Header $fixedHeaders -Delimiter ','

        foreach ($row in $data) {
            if ($row.net_amount)   { $row.net_amount   = [double]($row.net_amount   -replace ',', '.' -replace '[^\d\.\-]', '') }
            if ($row.vat_amount)   { $row.vat_amount   = [double]($row.vat_amount   -replace ',', '.' -replace '[^\d\.\-]', '') }
            if ($row.gross_amount) { $row.gross_amount = [double]($row.gross_amount -replace ',', '.' -replace '[^\d\.\-]', '') }
        }

        foreach ($row in $data) { $allData.Add($row) }
        Write-Host "  OK: ${filePath} ($($data.Count) Zeilen)"

    } catch {
        Write-Host "  Fehler bei ${filePath}: $_" -ForegroundColor Red
    }
}

if ($allData.Count -gt 0 -and $allData[0].PSObject.Properties.Name -contains "settlement_reference_time") {
    $allData = $allData | Sort-Object settlement_reference_time
}

Write-Host ""
Write-Host "Erstelle Excel mit $($allData.Count) Zeilen..." -ForegroundColor Cyan

$allData | Export-Excel -Path $excelOutputFile `
    -WorksheetName "CombinedData" `
    -AutoSize `
    -FreezeTopRow `
    -BoldTopRow `
    -NoNumberConversion "*"

$excel = Open-ExcelPackage -Path $excelOutputFile
$ws = $excel.Workbook.Worksheets["CombinedData"]
$lastRow = $ws.Dimension.End.Row

foreach ($col in @(5, 6, 7)) {
    Set-ExcelColumn -Worksheet $ws -Column $col -NumberFormat '#,##0.00;[Red]-#,##0.00'
}

Set-ExcelColumn -Worksheet $ws -Column 8 -NumberFormat 'DD.MM.YYYY HH:MM'

$minWidths = @{ 1=10; 2=20; 3=20; 4=8; 5=14; 6=14; 7=14; 8=22; 9=20 }
foreach ($col in $minWidths.Keys) {
    if ($ws.Column($col).Width -lt $minWidths[$col]) {
        $ws.Column($col).Width = $minWidths[$col]
    }
}

Add-ConditionalFormatting -Worksheet $ws `
    -Range "G2:G$lastRow" `
    -RuleType LessThan `
    -ConditionValue 0 `
    -BackgroundColor "FFD0D0" `
    -ForegroundColor "CC0000"

$ws.Cells["J1"].Value = "Gesamtsumme (gross_amount)"
$ws.Cells["J1"].Style.Font.Bold = $true
$ws.Cells["J2"].Formula = "=SUM(G2:G$lastRow)"
$ws.Cells["J2"].Style.Numberformat.Format = '#,##0.00;[Red]-#,##0.00'

Add-ExcelTable -Range $ws.Cells["A1:I$lastRow"] `
    -TableName "Transactions" `
    -TableStyle Medium9

Close-ExcelPackage $excel

Write-Host "Excel gespeichert: $excelOutputFile" -ForegroundColor Green
Write-Host "Zeilen gesamt: $($allData.Count)"
Write-Host ""

# ────────────────────────────────────────────────────────────────
# TEIL 3: Fertige Excel-Datei in den Zielordner verschieben
# ────────────────────────────────────────────────────────────────

if (-not [string]::IsNullOrWhiteSpace($TargetFolder)) {
    $targetFile = Join-Path $TargetFolder "alleRechnungen.xlsx"

    try {
        if (-not (Test-Path $TargetFolder)) {
            New-Item -Path $TargetFolder -ItemType Directory -Force | Out-Null
            Write-Host "Zielordner erstellt: $TargetFolder"
        }

        Move-Item -Path $excelOutputFile -Destination $targetFile -Force
        Write-Host "Excel verschoben nach: $targetFile" -ForegroundColor Green

    } catch {
        Write-Host "Fehler beim Verschieben: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Datei liegt noch unter: $excelOutputFile" -ForegroundColor Yellow
    }
} else {
    Write-Host "Kein Zielordner konfiguriert – Excel liegt unter: $excelOutputFile" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Fertig! Viel Erfolg mit der Auswertung." -ForegroundColor Green
