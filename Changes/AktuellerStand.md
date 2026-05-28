# Aktueller Stand — viafintech Zahllauf Dashboard

Stand: 2026-05-21
Branch: `claude/setup-payment-dashboard-nqvRw`
Letzter Commit: `43176c9 fix: Dispatcher.Invoke-Overload + Runspace-Scriptblock-Scope`

## Phasen-Status

- **Phase 1 — Skeleton:** fertig.
  WPF-GUI, Sidebar, Dark Theme, RichTextBox-Log, ISO-KW, 5 Schritte als Pages.
- **Phase 2 — Skript-Integration:** fertig.
  Schritt 2 (`RechnungenausMailinExcel.ps1`) und Schritt 4 (`viafintech_Task_Skript.ps1`) starten ueber `Start-RunspaceJob`.
- **Phase 3 — Config-GUI fuer Pfade:** fertig.
  `MsgFolder`, `TaskFolder`, `ZahllaufFolder` werden in der GUI per FolderBrowserDialog gepflegt und in `Config/default.json` gespeichert. Wrapper geben die Pfade an die unveraenderten Originalskripte (nur Param-Block ergaenzt) weiter.
- **Stabilisierung:** fertig (Commits `f84a89e`, `2b074f0`, `caf5c71`, `43176c9`).
- **UI-Texte ausgelagert (2026-05-28):** fertig.
  Statische GUI-Texte (H1/H2, Anleitungen, Button-/CheckBox-Labels, Sektions-Labels, `Window.Title`, Expander-Header) liegen jetzt zentral und editierbar in `Config/strings.de.json` (UTF-8). Neue `Functions/Strings.ps1` mit `Import-AppStrings` + Mapping-Tabelle `$Script:StringMap` + `Apply-Strings`; Aufruf in `Add_Loaded` VOR `Update-Context`. Bisher unbenannte Textelemente im XAML haben ein `x:Name` bekommen. Dynamische Laufzeit-Texte (KW/Datum/Bezeichnung/Ordner/Status/Counter) bleiben unveraendert in `Update-Context`/`OnComplete`. Bearbeiter-Doku: `doc/TEXTE_BEARBEITEN.md`.

## Funktional getestet

- Schritt 1: Pfad oeffnen, Schritt markieren — OK.
- Schritt 3: Bezeichnung in Clipboard, Prosos starten (`C:\Program Files (x86)\PROSOZ Herten\OPEN PROSOZ\Anwendungen\OpenStarter.exe /app OpenClient.exe`) — OK.
- Einstellungen: Pfade speichern und neu laden — OK.
- Automatisches Fehler-Log + Log-Datei-Button (Notepad) — OK.

## Noch offen / zu testen

- Schritt 2 End-to-End (`alleRechnungen.xlsx` muss am Ende existieren).
- Schritt 4 End-to-End nach Pull von `43176c9` (Dispatcher.Invoke-Fix).
- Schritt 5 (Abschluss).
- Persistenz der Einstellungen ueber App-Neustart.

## Behobene Bugs (Chronologie)

1. **Modul-Scope:** `Write-Log` aus `Functions/Logging.ps1` war in `.psm1`-Modulen unsichtbar.
   Fix: `function global:Write-Log`, ebenso `Clear-Log`, `Get-IsoCalendarWeek`, `Start-RunspaceJob`. Wichtige Variablen zusaetzlich als `$global:AppRoot/Config/LogBox/LogFile/Window` gespiegelt.
2. **Cmdlet-Lookup nach Runspace-Disposal:** PS5.1 verliert `Get-Date`, `Start-Process` etc. auf dem Main-Thread, wenn ein Runspace im WPF-Event-Handler disposed wird.
   Fix: `OnComplete` im DispatcherTimer-Tick **vor** `Dispose()` ausfuehren. Kritische Aufrufe durch reines .NET ersetzt (`[DateTime]::Now`, `[System.IO.File]::AppendAllText`, `[System.Diagnostics.Process]::Start`, `[System.IO.Directory]::Exists`, `[System.Windows.Clipboard]::SetText`).
3. **Dispatcher.Invoke-Overload-Verwirrung:** `$dispatcher.Invoke($action, @($a,$b,$c))` wurde von PS5.1 als `Invoke(Delegate, DispatcherPriority)` aufgeloest -> Cast-Error.
   Fix: `GetNewClosure()` + `[void]$dispatcher.Invoke([Action]$append)` ohne Args-Array.
4. **Scriptblock-SessionState-Affinitaet:** Ein Scriptblock, der einem Runspace uebergeben wird, behaelt seinen Ursprungs-SessionState und sieht die inneren `Write-Log`/`Write-Host`/`Read-Host` nicht.
   Fix: `$ScriptBlock.ToString()` uebergeben und im Runspace via `[scriptblock]::Create($UserScriptText)` neu erzeugen.

## Architektur (aktuell)

- `Main.ps1` laedt XAML, dotsourced `Functions/*.ps1`, importiert `Modules/*.psm1`, mirrort die Script-Variablen `AppRoot/Config/LogBox/LogFile/Window` ins `$global:`-Scope.
- `Functions/RunspaceHelpers.ps1` baut STA-Runspace, setzt Variablen via `SessionStateProxy.SetVariable`, definiert inner `Write-Log`/`Write-Host`/`Read-Host`, erzeugt den User-Scriptblock im Runspace neu, pollt per `DispatcherTimer` (250 ms), ruft `OnComplete` vor `Dispose`.
- `Functions/ConfigLoader.ps1` liefert `Import-AppConfig` und `Save-AppConfig` (Set-Content).
- Module rufen ihre Skripte ausschliesslich ueber `Start-RunspaceJob` auf und faerben den UI-Status (`BrushSuccess` / `BrushDanger`).
- Originalskripte in `Scripts/` haben nur einen Param-Block bekommen (`$SourcePath`, `$TargetBase` bzw. `$MsgFolder`, `$TargetFolder`, `$CsvOrdner`, `$NurExcel`).

## Dateien (Stand 43176c9)

```
Main.ps1
Config/default.json
Changes/AktuellerStand.md          <-- diese Datei
doc/CLAUDE.MD
doc/prompt.md
Config/strings.de.json             # statische UI-Texte (UTF-8) -> doc/TEXTE_BEARBEITEN.md
doc/TEXTE_BEARBEITEN.md            # Bearbeiter-Anleitung fuer die UI-Texte
Functions/
  ConfigLoader.ps1                 # Import-AppConfig, Save-AppConfig
  DateHelpers.ps1                  # global:Get-IsoCalendarWeek
  Logging.ps1                      # global:Write-Log, global:Clear-Log
  RunspaceHelpers.ps1              # global:Start-RunspaceJob
  Strings.ps1                      # Import-AppStrings, Apply-Strings, $Script:StringMap
Modules/
  Step1_Mail.psm1
  Step2_Rechnungen.psm1            # Invoke-ExtractRechnungen, Invoke-Step2Run
  Step3_Pruflauf.psm1              # Invoke-Step3Copy, Invoke-Step3OpenProsos
  Step4_Auszahlung.psm1            # Invoke-StartTaskScript, Invoke-Step4Run
  Step5_Abschluss.psm1
Scripts/                           # Originale, nur Param-Block ergaenzt
  RechnungenausMailinExcel.ps1
  viafintech_Task_Skript.ps1
Views/MainWindow.xaml
```
