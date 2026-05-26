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

## Aktueller Stand (Stand 2026-05-26)

- **Phase 1 (Skeleton):** fertig — WPF-GUI, Sidebar, Dark Theme, RichTextBox-Log, ISO-KW, 5 Schritte als Pages.
- **Phase 2 (Skript-Integration):** fertig — Originalskripte laufen via `Start-RunspaceJob`:
  - Schritt 2: `Scripts/RechnungenausMailinExcel.ps1` via `Invoke-ExtractRechnungen` (`-NurExcel`, `-CsvOrdner`, `-MsgFolder`, `-TaskFolder`). Excel wird nicht-interaktiv nach `TargetFolder` verschoben (frueherer `Read-Host`-Netzwerk-Prompt entfernt).
  - Schritt 4: `Scripts/viafintech_Task_Skript.ps1` via `Invoke-StartTaskScript` (`-TaskFolder`, `-ZahllaufFolder`). Bei existierendem Zielordner erscheint jetzt ein `Confirm-Dialog` (Ja/Nein) statt stillem Ueberspringen.
- **Phase 3 (Config-GUI):** fertig — `MsgFolder`, `TaskFolder`, `ZahllaufFolder` werden in der GUI per FolderBrowserDialog gepflegt, in `Config/default.json` gespeichert und ueber Wrapper an die Skripte uebergeben.
- **Phase 5 (Abschluss):** fertig —
  - `Invoke-Step5MailTemplate -Bezeichnung` ruft `Scripts/MailanHaushalt.ps1` via Runspace, oeffnet einen Outlook-Entwurf mit Betreff `<Bezeichnung> Prueflauf - steht zur weiteren Kontrolle bereit`. Die Bezeichnung wird aus `Main.ps1` (`$Script:Context.Bezeichnung`) uebergeben, **nicht** im Modul berechnet (`Get-ZahllaufContext` ist aus Modulen nicht erreichbar).
  - `Invoke-Step5Finish` raeumt auf: `.msg`-Dateien + `extracted_Attachements` loeschen, Tageslog nach `\\VJC\GB1_Systembetreuung\Taskplaner\<Bezeichnung>_<yyyyMMdd_HHmmss>_<guid8>.log` sichern. Gesteuert ueber `Behavior.SimulateCleanup` (Default `true`).
- **Simulation-Toggle:** CheckBox in den Einstellungen (`ChkSetSimulate`) -> `Config.Behavior.SimulateCleanup`. Bei `true` wird das Loeschen nur protokolliert, nicht ausgefuehrt.
- **Stabilisierung:** fertig — Modul-Scope (`global:`), Cmdlet-Lookup-Verlust nach Runspace-Disposal, Dispatcher.Invoke-Overload, Scriptblock-SessionState-Affinitaet behoben.
- **Getestet OK (lokal vom Maintainer):** Schritt 1, Schritt 3, Einstellungen, Fehler-Log, Schritt 5 Mail-Vorlage.
- **Offen:** Schritt 2/4/5 End-to-End final, Persistenz ueber App-Neustart final bestaetigen.

Detail-Log: siehe `Changes/AktuellerStand.md`.

## Architektur-Notizen

- **Globaler Scope statt Script-Scope:** Da Module (`.psm1`) den `$Script:`-Scope von `Main.ps1` nicht sehen, sind kritische Funktionen als `function global:...` deklariert (`Write-Log`, `Clear-Log`, `Get-IsoCalendarWeek`, `Start-RunspaceJob`) und Variablen werden gespiegelt (`$global:AppRoot`, `$global:Config`, `$global:LogBox`, `$global:LogFile`, `$global:Window`). Der `$Script:`-Prefix bleibt in `Main.ps1` Primaerquelle.
- **Runspaces:** STA-Runspace via `Start-RunspaceJob` (`Functions/RunspaceHelpers.ps1`).
  - Shared State: `[hashtable]::Synchronized(@{ LogBox; Dispatcher; LogFile; AppRoot; Done; Error; Result; DialogResult })`.
  - Inner Proxies: `Write-Log` (RichTextBox + Datei via .NET `AppendAllText`), `Write-Host` (Farbe -> Level), `Read-Host` (liefert `''`), `Confirm-Dialog`.
  - User-Scriptblock wird im Runspace via `[scriptblock]::Create($UserScriptText)` neu erzeugt, sonst behaelt er seinen Ursprungs-SessionState und sieht die Proxies nicht.
  - Abschluss-Erkennung via `DispatcherTimer` (250 ms). `OnComplete` laeuft **vor** `Dispose()`, um den PS5.1-Cmdlet-Lookup-Bug zu vermeiden.
- **Confirm-Dialog (Dialoge aus dem Runspace):** Die Proxy zeigt per `[void]$Sync.Dispatcher.Invoke([Action]$closure)` eine `[System.Windows.MessageBox]::Show(...)` auf dem UI-Thread und blockiert den Runspace-Thread bis zur Antwort. Ergebnis wird in `$Sync.DialogResult` abgelegt und als `[bool]` (Yes) zurueckgegeben. Wichtig: fuer Rueckgabewerte **nicht** `[Func[...]]`, sondern Ergebnis ueber `$Sync` + `[Action]`-Cast, sonst greift wieder die Overload-Falle. Scripts im Runspace rufen `Confirm-Dialog` einfach auf (Funktion liegt im Eltern-Scope).
- **Dispatcher.Invoke:** Immer `GetNewClosure()` + `[void]$dispatcher.Invoke([Action]$closure)` ohne Args-Array, sonst kollidiert PS5.1 mit der `DispatcherPriority`-Overload.
- **.NET statt Cmdlet:** In kritischen Pfaden (Logging, Process-Start, File-Existenz, Clipboard) reines .NET (`[System.IO.File]::AppendAllText`, `[System.Diagnostics.Process]::Start`, `[System.IO.Directory]::Exists`, `[System.Windows.Clipboard]::SetText`, `[DateTime]::Now`).
- **Encoding-Falle (STRIKT):** `.ps1`-Dateien **ohne BOM** (z. B. `RechnungenausMailinExcel.ps1`) liest PS5.1 als CP1252. Multibyte-UTF-8 in **String-Literalen** kann den Parser zerschiessen — En-Dash `–`/Em-Dash `—` enden auf Byte `0x93`/`0x94` = typografische Anfuehrungszeichen, die PowerShell als String-Delimiter behandelt. Regel: in BOM-losen Dateien **nur ASCII** in Code/Strings (Umlaute als `ue`/`ae`/`oe`, normaler Bindestrich statt En-Dash). Box-Zeichen in **Kommentaren** sind harmlos. Dateien **mit** BOM (`viafintech_Task_Skript.ps1`, `MailanHaushalt.ps1`) vertragen UTF-8.
- **Funktions-Sichtbarkeit aus Modulen:** Module sehen nur globalen Scope. Dot-gesourcte Helfer wie `Get-ZahllaufContext` (nur `Get-IsoCalendarWeek` ist `global:`) sind aus `.psm1` **nicht** erreichbar — berechnete Werte daher in `Main.ps1` ermitteln und als Parameter uebergeben.
- **Originalskripte in `Scripts/`** bekommen primaer einen Param-Block; punktuelle Anpassungen (interaktive Prompts -> nicht-interaktiv, `Confirm-Dialog`) sind erlaubt, Body sonst unveraendert.
- **Skript-Pfade** kommen aus `Config/default.json` (`Paths.ScriptRechnungen`, `Paths.ScriptTask`, `Paths.ScriptMail`); Arbeits-Pfade (`MsgFolder`, `TaskFolder`, `ZahllaufFolder`) und `Behavior.SimulateCleanup` ebenfalls dort und ueber GUI editierbar.
- **UI-Status** (`Txt2LastRun`, `Txt4Status`, `TxtSettingsStatus`) wird in `OnComplete` gruen (Brush `Success`) oder rot (Brush `Danger`) gefaerbt.

## Datei-Layout

```
Main.ps1                          # Einstieg, XAML, Event-Handler, Settings, $global:-Mirror
Config/default.json               # Paths.{Msg,Task,Zahllauf}Folder + Skriptnamen, Behavior.SimulateCleanup, UI-Defaults
Changes/AktuellerStand.md         # Detailliertes Stand- und Bug-Log
doc/CLAUDE.MD                     # diese Datei
doc/prompt.md
Functions/
  ConfigLoader.ps1                # Import-AppConfig, Save-AppConfig
  DateHelpers.ps1                 # global:Get-IsoCalendarWeek + Get-ZahllaufContext (script-scope)
  Logging.ps1                     # global:Write-Log, global:Clear-Log
  RunspaceHelpers.ps1             # global:Start-RunspaceJob + Inner-Proxies (inkl. Confirm-Dialog)
Modules/
  Step1_Mail.psm1                 # Invoke-Step1OpenPath, Invoke-Step1MarkDone
  Step2_Rechnungen.psm1           # Invoke-ExtractRechnungen, Invoke-Step2Run, Invoke-Step2OpenPath
  Step3_Pruflauf.psm1             # Invoke-Step3Copy, Invoke-Step3OpenProsos
  Step4_Auszahlung.psm1           # Invoke-StartTaskScript, Invoke-Step4Run, Invoke-Step4OpenTask
  Step5_Abschluss.psm1            # Invoke-Step5MailTemplate, Invoke-Step5Finish
Scripts/                          # Originale - Param-Block + punktuelle Anpassungen
  RechnungenausMailinExcel.ps1    # KEIN BOM -> nur ASCII in Strings!
  viafintech_Task_Skript.ps1      # UTF-8 mit BOM
  MailanHaushalt.ps1              # UTF-8 mit BOM; Outlook-Entwurf, -Bezeichnung-Param
Views/MainWindow.xaml
```