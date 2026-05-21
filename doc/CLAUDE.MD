# CLAUDE.md – Regeln für Claude Code (viafintech Zahllauf Dashboard)

**Du bist Claude Code für das Projekt "viafintech Zahllauf Dashboard".**  
Du arbeitest im Stil des bestehenden **JC-Dashboards** des Maintainers.

---

## Wichtige Regeln (STRIKT befolgen)

### 1. Immer zuerst lesen
- **Vor jeder Aufgabe** liest du zuerst die Datei `docs/CLAUDE.md` (diese Datei).
- Du hältst dich **genau** an alle hier genannten Regeln.

### 2. Projektspezifische Prinzipien
- **Single-Maintainer** – Halte alles so einfach und wartbar wie möglich.
- **PowerShell 5.1 kompatibel** – Keine PowerShell 7 Features, keine `#Requires -Version 7`.
- **`$Script:`-Prefix** – Alle script-weiten Variablen werden mit `$Script:` deklariert (z. B. `$Script:Window`, `$Script:Config`).
- **Keine unnötige Abstraktion** – Keine übertriebenen Klassen, Factorys oder Dependency Injection. Einfache Funktionen und direkte Aufrufe reichen.
- **Keine Emoji in PowerShell-Strings** – Weder in `Write-Host`, `Write-Log` noch in UI-Texten.
- **Token-Spar-Regel** – Sei präzise. Vermeide unnötige Erklärungen, lange Kommentare und redundanten Code.

### 3. Arbeitsweise bei Änderungen
- **Änderungen > 50 Zeilen:**  
  Schreibe **zuerst einen kurzen Plan** (max. 8 Sätze) und warte auf die Bestätigung des Maintainers, bevor du Code schreibst.
- **Nur die wirklich nötigen Dateien** lesen und ändern.
- Bevorzuge **bestehende Strukturen** erweitern statt neue Dateien anzulegen.

### 4. Ausgabe-Format (für jede Aufgabe)
Du antwortest **immer** in folgender Reihenfolge:

1. **Kurzer Plan** (max. 8 Sätze)
2. **Konkrete Code-Änderungen**  
   - Zeige **komplette neue Funktionen** oder **komplette geänderte Abschnitte**.
   - Gib **genau an**, wo der Code eingefügt oder ersetzt werden muss (z. B. "In `Modules/Step2_Rechnungen.psm1` nach Zeile 42 einfügen").
   - Verwende klare Marker wie:
     ```powershell
     # === NEU: Funktion Invoke-ExtractRechnungen ===
     function Invoke-ExtractRechnungen { ... }
     ```

---

## Aktueller Stand (Stand 2026-05-21, Commit `43176c9`)

- **Phase 1 (Skeleton):** fertig — WPF-GUI, Sidebar, Dark Theme, RichTextBox-Log, ISO-KW, 5 Schritte als Pages.
- **Phase 2 (Skript-Integration):** fertig — Originalskripte laufen via `Start-RunspaceJob`:
  - Schritt 2: `Scripts/RechnungenausMailinExcel.ps1` via `Invoke-ExtractRechnungen` (`-NurExcel`, `-CsvOrdner`, `-MsgFolder`, `-TaskFolder`).
  - Schritt 4: `Scripts/viafintech_Task_Skript.ps1` via `Invoke-StartTaskScript` (nicht-interaktiv, `-TaskFolder`, `-ZahllaufFolder`, prueft Zielordner unter `Zahllauf\`).
- **Phase 3 (Config-GUI):** fertig — `MsgFolder`, `TaskFolder`, `ZahllaufFolder` werden in der GUI per FolderBrowserDialog gepflegt, in `Config/default.json` gespeichert und ueber Wrapper an die Skripte uebergeben.
- **Stabilisierung:** fertig — Modul-Scope (`global:`), Cmdlet-Lookup-Verlust nach Runspace-Disposal, Dispatcher.Invoke-Overload, Scriptblock-SessionState-Affinitaet behoben.
- **Getestet OK:** Schritt 1 (Pfad oeffnen / markieren), Schritt 3 (Clipboard, Prosos), Einstellungen speichern/laden, Fehler-Log + Notepad-Button.
- **Offen:** Schritt 2 und Schritt 4 End-to-End nach Pull `43176c9`, Schritt 5, Persistenz ueber App-Neustart.

Detail-Log: siehe `Changes/AktuellerStand.md`.

## Architektur-Notizen

- **Globaler Scope statt Script-Scope:** Da Module (`.psm1`) den `$Script:`-Scope von `Main.ps1` nicht sehen, sind kritische Funktionen als `function global:...` deklariert (`Write-Log`, `Clear-Log`, `Get-IsoCalendarWeek`, `Start-RunspaceJob`) und Variablen werden gespiegelt (`$global:AppRoot`, `$global:Config`, `$global:LogBox`, `$global:LogFile`, `$global:Window`). Der `$Script:`-Prefix bleibt in `Main.ps1` Primaerquelle.
- **Runspaces:** STA-Runspace via `Start-RunspaceJob` (`Functions/RunspaceHelpers.ps1`).
  - Shared State: `[hashtable]::Synchronized(@{ LogBox; Dispatcher; LogFile; AppRoot; Done; Error; Result })`.
  - Inner Proxies: `Write-Log` (RichTextBox + Datei via .NET `AppendAllText`), `Write-Host` (Farbe -> Level), `Read-Host` (liefert `''`).
  - User-Scriptblock wird im Runspace via `[scriptblock]::Create($UserScriptText)` neu erzeugt, sonst behaelt er seinen Ursprungs-SessionState und sieht die Proxies nicht.
  - Abschluss-Erkennung via `DispatcherTimer` (250 ms). `OnComplete` laeuft **vor** `Dispose()`, um den PS5.1-Cmdlet-Lookup-Bug zu vermeiden.
- **Dispatcher.Invoke:** Immer `GetNewClosure()` + `[void]$dispatcher.Invoke([Action]$closure)` ohne Args-Array, sonst kollidiert PS5.1 mit der `DispatcherPriority`-Overload.
- **.NET statt Cmdlet:** In kritischen Pfaden (Logging, Process-Start, File-Existenz, Clipboard) reines .NET (`[System.IO.File]::AppendAllText`, `[System.Diagnostics.Process]::Start`, `[System.IO.Directory]::Exists`, `[System.Windows.Clipboard]::SetText`, `[DateTime]::Now`).
- **Originalskripte in `Scripts/`** bekommen nur einen Param-Block — Body bleibt unveraendert.
- **Skript-Pfade** kommen aus `Config/default.json` (`Paths.ScriptRechnungen`, `Paths.ScriptTask`); Arbeits-Pfade (`MsgFolder`, `TaskFolder`, `ZahllaufFolder`) ebenfalls dort und ueber GUI editierbar.
- **UI-Status** (`Txt2LastRun`, `Txt4Status`, `TxtSettingsStatus`) wird in `OnComplete` gruen (Brush `Success`) oder rot (Brush `Danger`) gefaerbt.

## Datei-Layout

```
Main.ps1                          # Einstieg, XAML, Event-Handler, Settings, $global:-Mirror
Config/default.json               # Paths.{Msg,Task,Zahllauf}Folder, Skriptnamen, UI-Defaults
Changes/AktuellerStand.md         # Detailliertes Stand- und Bug-Log
doc/CLAUDE.MD                     # diese Datei
doc/prompt.md
Functions/
  ConfigLoader.ps1                # Import-AppConfig, Save-AppConfig
  DateHelpers.ps1                 # global:Get-IsoCalendarWeek
  Logging.ps1                     # global:Write-Log, global:Clear-Log
  RunspaceHelpers.ps1             # global:Start-RunspaceJob
Modules/
  Step1_Mail.psm1
  Step2_Rechnungen.psm1           # Invoke-ExtractRechnungen, Invoke-Step2Run
  Step3_Pruflauf.psm1             # Invoke-Step3Copy, Invoke-Step3OpenProsos
  Step4_Auszahlung.psm1           # Invoke-StartTaskScript, Invoke-Step4Run
  Step5_Abschluss.psm1
Scripts/                          # Originale - nur Param-Block ergaenzt
  RechnungenausMailinExcel.ps1
  viafintech_Task_Skript.ps1
Views/MainWindow.xaml
```