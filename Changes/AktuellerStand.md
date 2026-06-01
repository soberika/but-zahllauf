# Aktueller Stand — viafintech Zahllauf Dashboard

Stand: 2026-06-01
Branch: `claude/eager-dijkstra-G7Vxx`
Letzter Commit: `4d74e47 Status-Banner + KW-Wechsel-Hinweis und Reset pro KW`

## Phasen-Status

- **Phase 1 — Skeleton:** fertig.
  WPF-GUI, Sidebar, Dark Theme, RichTextBox-Log, ISO-KW, 5 Schritte als Pages.
- **Phase 2 — Skript-Integration:** fertig.
  Schritt 2 (`RechnungenausMailinExcel.ps1`) und Schritt 4 (`viafintech_Task_Skript.ps1`) starten ueber `Start-RunspaceJob`.
- **Phase 3 — Config-GUI fuer Pfade:** fertig.
  `MsgFolder`, `TaskFolder`, `ZahllaufFolder` per FolderBrowserDialog, in `Config/default.json` gespeichert; `TempRoot`/`ExtractedFolder` werden via `Resolve-DerivedPaths` aus `MsgFolder` abgeleitet.
- **Phase 5 — Abschluss:** fertig.
  Mail-Vorlage (`MailanHaushalt.ps1`, `-Bezeichnung`/`-Summe`/`-ZielpfadLink`), Temp-Loeschen (respektiert `Behavior.SimulateCleanup`), Log-Sicherung in den Taskplaner-Ordner.
- **Aufgabenplanung (on demand):** fertig.
  `Modules/Step4_ScheduledTasks.psm1` stoesst registrierte Windows-Tasks an, Status nur ueber `Get-ScheduledTaskInfo`. Keine Credentials/Args/Task-XML im Repo.
- **UI-Texte ausgelagert (2026-05-28):** fertig.
  Statische Texte in `Config/strings.de.json`, `Functions/Strings.ps1` (`Import-AppStrings`, `$Script:StringMap`, `Apply-Strings`). Bearbeiter-Doku: `doc/TEXTE_BEARBEITEN.md`.
- **Schritt-Ueberarbeitung (2026-06-01):** fertig.
  Aufgabenorientierte Buttons je Schritt, PDF-Merge (Schritt 2, itextsharp), Gesamtsumme J2 (Schritt 3), Abgleich-Buttons (Schritt 4), Schritt 5 in zwei Aktionen. Rechtes "Globale Parameter"-Panel entfernt.

## GUI-Usability fuer Fachfremde (2026-06-01)

Erweiterungen, damit die Oberflaeche auch ohne Prozesswissen bedienbar ist:

1. **Sichtbarer Fortschritt (#1):** Erledigte Schritte werden in der Sidebar gruen,
   der aktive hell, offene gedaempft (`Update-StepProgress`). Top-Bar-Karte
   "Fortschritt" mit Zaehler "x / 5 erledigt"; dafuer wurden die Karten
   "Kalenderwoche" und "Jahr" aus der Top-Bar entfernt. "Schritt als erledigt
   markieren" (+ "Zahllauf abgeschlossen") rufen `Set-StepDone`: Status setzen,
   **pro KW** in `Config/state.json` persistieren (gitignored,
   `Functions/StepState.ps1`), zum naechsten Schritt springen. `Load-StepDone`
   laedt den Status beim KW-Wechsel.
2. **Bereitschaftsanzeige Schritt 2 (#2):** `Update-Readiness` zeigt beim Oeffnen
   die `.msg`-Lage (`Txt2Ready`); bei 0 Dateien / nicht erreichbarem Ordner
   orange Warnung + `Btn2Run` gesperrt. `BtnPrimary` hat dafuer einen
   `IsEnabled=False`-Trigger bekommen (gesperrte Primaer-Buttons werden ausgegraut).
3. **Pfad-Check beim Start (#4):** `Test-ConfigReady` prueft Msg-/Task-/Zahllauf-Ordner;
   bei fehlenden Ordnern Log-Warnung, MessageBox und Sprung in die Einstellungen.
4. **Status-Banner (#5):** Auto-Zeile zwischen Top-Bar und Body (`StatusBanner`).
   `global:Set-AppStatus`/`global:Hide-AppStatus` in `Functions/Logging.ps1`,
   Dispatcher-sicher. Zentral eingeklinkt statt in jedem Modul: `Write-Log -Level
   Error` und `Start-RunspaceJob` bei `$sync.Error`. Body-Grid jetzt `Grid.Row=2`,
   Log `Grid.Row=3`.
5. **KW-Wechsel-Hinweis + Reset (#6):** Beim KW-Wechsel Info-Banner
   "Bearbeitete KW: <Bezeichnung> - x/5 erledigt" (Start unterdrueckt via
   `$Script:SuppressWeekBanner`). Settings-Karte "Fortschritt" mit `BtnResetWeek`
   (Erledigt-Status der aktuellen KW zuruecksetzen).

**Zurueckgestellt:** #3 Tooltips fuer Fachbegriffe — Konzept steht (datengetrieben
aus `strings.de.json` via `Apply-Tooltips`), Umsetzung spaeter.

## Funktional getestet (lokal vom Maintainer)

- Schritt 1: Pfad/Outlook oeffnen, Schritt markieren — OK.
- Schritt 3: Bezeichnung in Clipboard, Prosos starten — OK.
- Einstellungen: Pfade speichern und neu laden — OK.
- Automatisches Fehler-Log + Log-Datei-Button (Notepad) — OK.
- Schritt 5: Mail-Vorlage (Outlook-Entwurf) — OK.

## Noch offen / zu testen

- Schritt 2 (Extraktion + PDF-Merge) und Schritt 4/5 End-to-End final.
- Persistenz (Einstellungen + Schritt-Status `state.json`) ueber App-Neustart final bestaetigen.
- Neue Usability-Features (#1/#2/#4/#5/#6) im laufenden Betrieb pruefen
  (in der Ausfuehrungsumgebung steht kein PowerShell/WPF zur Verfuegung, daher
  nur statische Validierung: JSON/XAML well-formed, `.ps1` BOM-los + ASCII,
  Klammern balanciert).
- `itextsharp.dll` einmalig unter `Assets\lib\` ablegen + entsperren (`Assets\lib\README.md`).

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

- `Main.ps1` laedt XAML, dotsourced `Functions/*.ps1`, importiert `Modules/*.psm1`, mirrort die Script-Variablen `AppRoot/Config/LogBox/LogFile/Window/StatusBanner/StatusBannerText` ins `$global:`-Scope.
- `Functions/RunspaceHelpers.ps1` baut STA-Runspace, definiert inner `Write-Log`/`Write-Host`/`Read-Host`/`Confirm-Dialog`, erzeugt den User-Scriptblock im Runspace neu, pollt per `DispatcherTimer` (250 ms), ruft `OnComplete` vor `Dispose`; bei `$sync.Error` wird das Status-Banner gezeigt.
- `Functions/Logging.ps1` liefert `global:Write-Log` (RichTextBox + Tageslog, Banner bei Level Error) sowie `global:Set-AppStatus`/`global:Hide-AppStatus`.
- `Functions/Strings.ps1` liefert `Import-AppStrings`, `$Script:StringMap`, `Apply-Strings` (statische UI-Texte aus `strings.de.json`).
- `Functions/StepState.ps1` liefert `Import-StepState`/`Save-StepState` (Erledigt-Status pro KW in `Config/state.json`, gitignored).
- `Functions/ConfigLoader.ps1` liefert `Import-AppConfig`, `Save-AppConfig`, `Resolve-DerivedPaths`.
- Module rufen ihre Skripte ausschliesslich ueber `Start-RunspaceJob` auf und faerben den UI-Status (`BrushSuccess` / `BrushDanger`).
- Originalskripte in `Scripts/` haben nur einen Param-Block + punktuelle nicht-interaktive Anpassungen bekommen.

## Dateien (Stand 4d74e47)

```
Main.ps1
Config/default.json                # Pfade, Skriptnamen, Behavior, ScheduledTasks, Hints, UI
Config/strings.de.json             # statische UI-Texte (UTF-8) -> doc/TEXTE_BEARBEITEN.md
Config/state.json                  # Erledigt-Status pro KW (gitignored, zur Laufzeit)
Changes/AktuellerStand.md          <-- diese Datei
doc/CLAUDE.MD                      # Regeln + Architektur + Stand
doc/TEXTE_BEARBEITEN.md            # Bearbeiter-Anleitung fuer die UI-Texte
doc/prompt.md
Functions/
  ConfigLoader.ps1                 # Import-AppConfig, Save-AppConfig, Resolve-DerivedPaths
  DateHelpers.ps1                  # global:Get-IsoCalendarWeek + Get-ZahllaufContext
  Logging.ps1                      # global:Write-Log, Clear-Log, Set-AppStatus, Hide-AppStatus
  RunspaceHelpers.ps1              # global:Start-RunspaceJob + Inner-Proxies
  Strings.ps1                      # Import-AppStrings, Apply-Strings, $Script:StringMap
  StepState.ps1                    # Import-StepState, Save-StepState (state.json)
Modules/
  Step1_Mail.psm1
  Step2_Rechnungen.psm1            # Extraktion, PDF-Merge, Datei-Oeffner
  Step3_Pruflauf.psm1              # Copy, ReadSumme (J2), OpenProsos
  Step4_Auszahlung.psm1            # StartTaskScript, Abgleich-Dateien
  Step4_ScheduledTasks.psm1        # Windows-Tasks on demand + Status
  Step5_Abschluss.psm1             # Mail-Vorlage + Abschluss
Scripts/                           # Originale, Param-Block + punktuelle Anpassungen
  RechnungenausMailinExcel.ps1     # KEIN BOM -> nur ASCII in Strings
  viafintech_Task_Skript.ps1       # UTF-8 mit BOM
  MailanHaushalt.ps1               # UTF-8 mit BOM; Outlook-Entwurf
Views/MainWindow.xaml
```
