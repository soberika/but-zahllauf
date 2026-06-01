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

## Aktueller Stand (Stand 2026-06-01)

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
- **KW-Auswahl:** fertig — `Get-ZahllaufContext` rechnet mit einem Stichtag (`-RefDate`, Default = letzter Sonntag der vergangenen Woche). Neuer Helfer `Get-RecentWeeks` liefert die letzten 4 Wochen. Top-Bar zeigt statt eines Refresh-Buttons eine ComboBox `CmbWeek`; Auswahl loest `Update-Context` aus. Button "Werte neu berechnen" entfernt.
- **Aufgabenplanung (on demand):** fertig — neues Modul `Modules/Step4_ScheduledTasks.psm1`. `Invoke-StartScheduledTask` stoesst einen in der Windows-Aufgabenplanung registrierten Task im Runspace an (`Start-ScheduledTask`, Fallback `schtasks.exe /Run /TN "<Pfad><Name>"`), pollt den Abschluss **ausschliesslich** ueber `Get-ScheduledTaskInfo` (`LastTaskResult`/`LastRunTime`; `0x41301` = laeuft, danach neue `LastRunTime`) und faerbt `Txt4TaskStatus` gruen/rot. `Confirm-Dialog` vor jedem Start. Wrapper `Invoke-ScheduledTaskById` loest den Task per `Id` aus `Config.ScheduledTasks` auf (analog `Invoke-Step4Run`). Buttons werden in `Initialize-ScheduledTaskButtons` (Main.ps1) dynamisch aus `Config.ScheduledTasks` in der Karte "Aufgabenplanung" (Schritt 4) erzeugt. **Sicherheit (STRIKT):** keine Credentials, Argumente oder Task-XML im Repo, in der Config oder im Log; das Dashboard referenziert ausschliesslich `TaskPath`+`TaskName` und loggt nur den Ergebnis-Code. Niemals `schtasks /Query /XML`.
- **UI-Ueberarbeitung (2026-05-27):**
  - UI-Texte auf echte Umlaute umgestellt (XAML wird mit `-Encoding UTF8` geladen, daher gefahrlos; ASCII-Regel gilt nur fuer BOM-lose `.ps1`). Button-Labels aufgabenorientiert statt Skriptnamen ("Rechnungen extrahieren", "Auszahlungslauf starten").
  - Schritt 2: nur noch **eine** Option-Checkbox `Chk2NurExcel`. Die fruehere zweite Checkbox `Chk2NurExcelCsv` war funktional identisch (kein CSV-Ordner-Feld, `-CsvOrdner` nie uebergeben) und wurde entfernt; `Invoke-Step2Run` reicht `-NurExcel` direkt durch.
  - Schritt 3: Checkliste verdrahtet (`Chk3Done0..3`) und **pro KW** in `Config.Checklist[<OrdnerName>]` persistiert. `Sync-Step3Checklist` laedt beim KW-Wechsel, `Save-Step3Checklist` speichert pro Klick; Zaehler `Txt3Checklist` ("x/4 erledigt", gruen bei vollstaendig).
  - Schritt 4: Platzhalter-Button "Letzte Korrektur pruefen" (`Btn4Check`) + Funktion `Invoke-Step4CheckCorrection` ersatzlos entfernt (war ein reiner Log-No-op).
  - **Bebilderte Hinweise (Galerie):** pro Step ein `Expander` "Bebilderte Hinweise" mit `WrapPanel` (`Sp1Hints..Sp5Hints`), datengetrieben aus `Config.Hints.Step1..Step5` (`{ Image, Caption }`). `Initialize-HintGalleries` baut Thumbnail-Kacheln; fehlt die Bilddatei, erscheint eine Platzhalter-Kachel ("Bild folgt"). Klick oeffnet die Grossansicht via `global:Show-HintImage` (modal). Closure-Werte (Pfad/Caption/Owner) **vor** `GetNewClosure()` in lokale Variablen gehoben (siehe Closure-Scope-Note). Bilder liegen unter `Assets/hints/` und sind **gitignored** (Screenshots interner Systeme nicht committen).
- **UI-Texte ausgelagert (2026-05-28):** Statische GUI-Texte liegen zentral und editierbar in `Config/strings.de.json` (UTF-8, echte Umlaute erlaubt). Neue `Functions/Strings.ps1` mit `Import-AppStrings` (analog `Import-AppConfig`), Mapping-Tabelle `$Script:StringMap` (Element-`x:Name` -> Property `Text`/`Content`/`Header` -> Schluessel) und `Apply-Strings`; Aufruf in `Add_Loaded` **vor** `Initialize-WeekPicker`/`Update-Context`. Bisher unbenannte In-Scope-Textelemente im XAML haben ein `x:Name` erhalten; bestehende Namen unveraendert. Fehlt ein Schluessel/Element, bleibt der XAML-Wert stehen + Warnung im Log. Dynamische Laufzeit-Texte (KW/Datum/Bezeichnung/Ordner/Status/Counter) bleiben **out of scope** in `Update-Context`/`OnComplete`; Task-Labels und Hinweis-Captions kommen weiter aus `Config/default.json`. Bearbeiter-Doku: `doc/TEXTE_BEARBEITEN.md`.
- **Schritt-Ueberarbeitung (2026-06-01):** Ergaenzung zum Stand 2026-05-28. Alle 5 Schritt-Karten bekamen aufgabenorientierte Buttons und je einen "Schritt als erledigt markieren"-Button (`Invoke-StepXMarkDone`, reines Log-Success).
  - **Schritt 1:** zwei getrennte Karten "Outlook" (`Btn1OpenOutlook` -> `Invoke-Step1OpenOutlook`) und "Speichern" mit **zwei Checkboxen mit eigenen Ordner-Buttons** (`Chk1Task` -> `Btn1OpenTask` = `TaskFolder`, `Chk1Mails` -> `Btn1OpenPath` = `MsgFolder`). Neuer Button "Pfad in Zwischenablage" (`Btn1CopyPath` -> `Invoke-Step1CopyPath`, kopiert `MsgFolder`). `Btn1OpenPath` zeigt jetzt auf `MsgFolder` statt `TempRoot`.
  - **Schritt 2:** neuer Excel-Oeffnen-Button (`Btn2OpenExcel`, `alleRechnungen.xlsx` im `TaskFolder`); `Btn2OpenPath` zeigt jetzt auf `ExtractedFolder`. **PDF-Zusammenfassung (neu):** Karte "PDFs zusammenfassen" -> `Invoke-Step2MergePdf` fasst die extrahierten Einzel-PDFs (`ExtractedFolder`, sortiert nach `CreationTime` = Mail-/Extraktionsreihenfolge) zu `alleRechnungen_Anhaenge.pdf` im `TaskFolder` zusammen. Nutzt **`itextsharp.dll` aus `Assets\lib`** (reines .NET, kein Install/Internet) via `PdfCopy`/seitenweisem Kopieren, **kompatibel zu iTextSharp 4.x UND 5.x**. Abhaengigkeiten (z. B. `BouncyCastle.Crypto.dll` bei 5.x) werden vor `itextsharp.dll` aus dem Ordner geladen; fehlende Abhaengigkeit wird exakt aus der `ReflectionTypeLoadException` gemeldet. Fehler beim Lesen einzelner PDFs werden gemeldet statt vom `Close()` verdeckt; bei 0 uebernommenen Seiten wird die leere Ausgabe wieder geloescht. Fehlt die DLL, bricht der Schritt sauber ab (nichts wird geloescht/ueberschrieben). DLL ist **gitignored**; Beschaffung + `Unblock-File` (gegen Fehler `0x80131515`) in `Assets\lib\README.md`. Buttons "Zusammengefasste PDF oeffnen" (`Btn2OpenMergedPdf`) und "Erledigt" (`Btn2Done`).
  - **Schritt 3:** die **pro-KW-Checkliste wurde komplett entfernt** (`Chk3Done0..3`, `Txt3Checklist`, `Config.Checklist`, Helfer `Sync-/Save-Step3Checklist`, `Get-ChecklistBoxes`, `Update-ChecklistCounter`; auch der Aufruf in `Update-Context`). Stattdessen **Gesamtsumme-Button** (`Btn3Summe` -> `Invoke-Step3ReadSumme`): liest Zelle **J2** aus `alleRechnungen.xlsx` per Excel-COM im Runspace, formatiert als `N2` (de-DE, Euro) und zeigt sie in `Txt3Summe`. Neuer "Erledigt"-Button (`Btn3Done`).
  - **Schritt 4:** Karte **"Abgleich der Gesamtsummen"** mit `Btn4OpenXlsx`/`Btn4OpenPdf` (`Invoke-Step4OpenAbgleichFile -Folder -Pattern`, Patterns `*alleRechnungen.xlsx` und `*ZahllisteHHSTGesamtBetr.pdf`). Die Buttons sind erst **nach erfolgreichem Lauf** aktiv; `Invoke-StartTaskScript`/`OnComplete` aktiviert sie und hinterlegt den Zielordner an `Button.Tag`. Erstellungs-Button von "Zahlliste-PDF oeffnen" zu **"Haushaltsstellengesamtbetraege oeffnen"** umbenannt; Karten-Ueberschrift `Txt4ErstellungTitle` ergaenzt. Neuer "Erledigt"-Button (`Btn4Done`).
  - **Schritt 5 in zwei Aktionen aufgeteilt:** `Invoke-Step5DeleteTemp` (`Btn5DeleteTemp`) loescht nur die temporaeren Dateien (`.msg` + `ExtractedFolder`-Anhaenge), respektiert `Behavior.SimulateCleanup`. `Invoke-Step5Finish` sichert **nur noch** das Tageslog nach `\\VJC\GB1_Systembetreuung\Taskplaner\...` und protokolliert die Uebergabe (kein Loeschen mehr darin); bekommt `-Bezeichnung` aus `Main.ps1`. Hinweistext `Txt5FinishNote`. **Mail-Vorlage erweitert:** `Invoke-Step5MailTemplate -Summe -ZielpfadLink`; `MailanHaushalt.ps1` (`-Summe`, `-ZielpfadLink`) baut den HTML-Body mit Gesamtsumme und einem `file://`-Link auf den Zielordner. Summe kommt aus `Txt3Summe`, Zielpfad = `ZahllaufFolder\OrdnerName`. **Fix:** `Get-ZahllaufContext`-Aufruf aus dem Step5-Modul entfernt (aus `.psm1` nicht erreichbar - siehe Architektur-Note).
  - **Abgeleitete Pfade konsistent zu den Einstellungen:** `Paths.TempRoot`/`Paths.ZahllaufRoot` aus `Config/default.json` entfernt. Neue Funktion `Resolve-DerivedPaths` (`Functions/ConfigLoader.ps1`) leitet `TempRoot = Split-Path -Parent MsgFolder` und `ExtractedFolder = TempRoot\extracted_attachments` ab; Aufruf **nach** `Import-AppConfig` und **nach** jedem Speichern der Einstellungen, sodass die Ableitungen einem geaenderten `MsgFolder` folgen. Einstellungen-Seite: Ordner-Label konkretisiert, Einleitungstext angepasst.
  - **Strings:** `Config/strings.de.json` an die neuen Elemente angepasst (neue Schluessel u. a. `Step1.Outlook*/Save*`, `Step2.Summary*/Merge*`, `Step3.Summe*`, `Step4.Erstellung*/Abgleich*`, `Step5.FinishNote` plus zahlreiche neue Button-Keys; Anleitungs- und Checklist-Keys entfernt).
  - **Rechte Spalte "Globale Parameter" entfernt:** Das dritte Body-Panel (`Grid.Column=2`, `LblRight*`/`Txt*Right`) war redundant zu Top-Bar und Schritt-Karten. Body-Grid jetzt 2-spaltig (`240/*`). `Jahr` als vierte Top-Bar-Karte ergaenzt (`TxtYear`/`LblTopYear`, `Update-Context` setzt `$ctx.Year`). Strings: `Right.*`-Schluessel raus, `Top.Year` rein; StringMap analog (`LblRight*` entfernt, `LblTopYear` ergaenzt).
- **Sichtbarer Fortschritt (2026-06-01):** Erledigte Schritte werden in der Sidebar gruen, offene gedaempft, der aktive Schritt hell gefaerbt (`Update-StepProgress`). Neue Top-Bar-Karte "Fortschritt" (`TxtProgress`, Zaehler "x / 5 erledigt"); dafuer wurden die Karten **"Kalenderwoche" (`TxtKW`/`LblTopKW`)** und **"Jahr" (`TxtYear`/`LblTopYear`)** aus der Top-Bar entfernt (inkl. `$ui`-Refs, `Update-Context`-Zuweisungen, StringMap, Strings `Top.KW`/`Top.Year` -> neu `Top.Progress`). Die "Schritt als erledigt markieren"-Buttons (+ `Btn5Finish`) rufen jetzt zusaetzlich `Set-StepDone <n>`: setzt den Status, persistiert **pro KW-Ordnername** in neuer `Config/state.json` (gitignored) via `Functions/StepState.ps1` (`Import-StepState`/`Save-StepState`) und springt zum naechsten Schritt. `Load-StepDone` laedt den Status beim KW-Wechsel (am Ende von `Update-Context`); `Import-StepState` in `Add_Loaded` vor `Initialize-WeekPicker`.
- **Bereitschaftsanzeige Schritt 2 (2026-06-01):** Neue Zeile `Txt2Ready` prueft beim Oeffnen von Schritt 2 (`Update-Readiness`, aufgerufen in `Show-Step`) den `MsgFolder`: bei 0 `.msg`-Dateien oder nicht erreichbarem Ordner orange Warnung + `Btn2Run` gesperrt, sonst gruene Meldung "N bereit". `BtnPrimary` hat dafuer einen `IsEnabled=False`-Trigger bekommen (gesperrte Primaer-Buttons werden jetzt ausgegraut wie `BtnBase`). Text ist dynamisch (kein strings.de.json-Schluessel).
- **Pfad-Check beim Start (2026-06-01):** `Test-ConfigReady` prueft in `Add_Loaded` (nach `Show-Step 1`) die drei Ordner `MsgFolder`/`TaskFolder`/`ZahllaufFolder` per `[System.IO.Directory]::Exists`. Sind welche leer/nicht erreichbar: Warnung ins Log, `MessageBox` mit Auflistung und automatischer Sprung in die Einstellungen (`Show-Step 6`). Sonst still.
- **Status-Banner (2026-06-01):** Neue Auto-Zeile zwischen Top-Bar und Body (`StatusBanner`/`TxtStatusBanner` + Button `BtnStatusDismiss` "Ausblenden"); Body-Grid jetzt `Grid.Row=2`, Log `Grid.Row=3`. `global:Set-AppStatus -Message -Level` (Info/Success/Warning/Error) faerbt + zeigt das Banner Dispatcher-sicher, `global:Hide-AppStatus` blendet aus (beide in `Functions/Logging.ps1`, `$global:StatusBanner`/`$global:StatusBannerText` in Main.ps1 gespiegelt). Eingeklinkt an zwei zentralen Stellen statt in jedem Modul: `global:Write-Log` ruft bei Level `Error` automatisch `Set-AppStatus`; `Start-RunspaceJob` zeigt im Timer-Tick bei `$sync.Error` das Banner (Runspace-Exceptions). Die **innere** Runspace-`Write-Log` ist davon unabhaengig (loggt nur), daher der zusaetzliche `$sync.Error`-Hook.
- **KW-Wechsel-Hinweis + Reset (2026-06-01):** Beim Wechsel der bearbeiteten KW zeigt `Update-Context` ein Info-Banner "Bearbeitete KW: <Bezeichnung> - x/5 erledigt" (unterdrueckt beim Start via `$Script:SuppressWeekBanner`, nach `Initialize-WeekPicker` freigegeben). Neue Settings-Karte "Fortschritt" mit `BtnResetWeek`: setzt `$Script:StepDone` zurueck, entfernt den KW-Key aus `Config/state.json` und aktualisiert die Anzeige. Strings: `Settings.ProgressTitle/Desc`, `Btn.Settings.ResetWeek`, `Btn.Status.Dismiss` (+ StringMap).
- **Getestet OK (lokal vom Maintainer):** Schritt 1, Schritt 3, Einstellungen, Fehler-Log, Schritt 5 Mail-Vorlage.
- **Offen:** Schritt 2 (Extraktion + PDF-Merge) und Schritt 4/5 End-to-End final, Persistenz ueber App-Neustart final bestaetigen. `itextsharp.dll` muss einmalig unter `Assets\lib\` abgelegt werden (siehe `Assets\lib\README.md`).

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
- **Closure-Scope bei dynamischen Handlern:** `GetNewClosure()` erzeugt fuer den Scriptblock einen neuen Modulscope. Einfache Variablen (`$ui`, lokale) werden per Wert eingefangen, aber explizite `$Script:`-Referenzen (`$Script:Window`) loesen im Closure auf den **leeren** neuen Scope auf (`-> $null`, `$null.Resources['Success']` = "Index auf NULL-Array"). Werte wie Brushes/Styles/StatusBlock daher **vor** dem Closure in lokale Variablen heben (siehe `Initialize-ScheduledTaskButtons` in `Main.ps1`).
- **.NET statt Cmdlet:** In kritischen Pfaden (Logging, Process-Start, File-Existenz, Clipboard) reines .NET (`[System.IO.File]::AppendAllText`, `[System.Diagnostics.Process]::Start`, `[System.IO.Directory]::Exists`, `[System.Windows.Clipboard]::SetText`, `[DateTime]::Now`).
- **Encoding-Falle (STRIKT):** `.ps1`-Dateien **ohne BOM** (z. B. `RechnungenausMailinExcel.ps1`) liest PS5.1 als CP1252. Multibyte-UTF-8 in **String-Literalen** kann den Parser zerschiessen — En-Dash `–`/Em-Dash `—` enden auf Byte `0x93`/`0x94` = typografische Anfuehrungszeichen, die PowerShell als String-Delimiter behandelt. Regel: in BOM-losen Dateien **nur ASCII** in Code/Strings (Umlaute als `ue`/`ae`/`oe`, normaler Bindestrich statt En-Dash). Box-Zeichen in **Kommentaren** sind harmlos. Dateien **mit** BOM (`viafintech_Task_Skript.ps1`, `MailanHaushalt.ps1`) vertragen UTF-8.
- **KW-Logik:** Abgerechnet wird die vergangene Woche. `Get-ZahllaufContext -RefDate <Sonntag>` leitet KW, Jahr, Bezeichnung, OrdnerName und Faelligkeit aus dem Stichtag ab; `Now/Date/Weekday` zeigen weiterhin das Tagesdatum. `Get-RecentWeeks -Count 4` baut die ComboBox-Liste (Sonntage der letzten 4 Wochen). Die Liste wird nur beim App-Start befuellt; nach einem Tageswechsel ist ein Neustart noetig.
- **Funktions-Sichtbarkeit aus Modulen:** Module sehen nur globalen Scope. Dot-gesourcte Helfer wie `Get-ZahllaufContext` (nur `Get-IsoCalendarWeek` ist `global:`) sind aus `.psm1` **nicht** erreichbar — berechnete Werte daher in `Main.ps1` ermitteln und als Parameter uebergeben.
- **Originalskripte in `Scripts/`** bekommen primaer einen Param-Block; punktuelle Anpassungen (interaktive Prompts -> nicht-interaktiv, `Confirm-Dialog`) sind erlaubt, Body sonst unveraendert.
- **Skript-Pfade** kommen aus `Config/default.json` (`Paths.ScriptRechnungen`, `Paths.ScriptTask`, `Paths.ScriptMail`); Arbeits-Pfade (`MsgFolder`, `TaskFolder`, `ZahllaufFolder`) und `Behavior.SimulateCleanup` ebenfalls dort und ueber GUI editierbar.
- **UI-Status** (`Txt2LastRun`, `Txt4Status`, `TxtSettingsStatus`) wird in `OnComplete` gruen (Brush `Success`) oder rot (Brush `Danger`) gefaerbt.
- **Abgeleitete Pfade:** `TempRoot`/`ExtractedFolder` stehen **nicht** mehr in `default.json`, sondern werden von `Resolve-DerivedPaths` aus `MsgFolder` berechnet (`TempRoot` = Elternordner, `ExtractedFolder` = `TempRoot\extracted_attachments`). Aufruf nach `Import-AppConfig` und nach jedem Settings-Save, damit sie einem geaenderten `MsgFolder` folgen.
- **PDF-Merge (Assets\lib, .NET-only):** `Invoke-Step2MergePdf` laedt `itextsharp.dll` zur Laufzeit per `Add-Type` aus `Assets\lib` (keine Installation/Internet/Admin). Vorher werden alle uebrigen DLLs des Ordners geladen (Abhaengigkeiten wie `BouncyCastle.Crypto.dll` bei 5.x). Seitenweises Kopieren via `PdfCopy` ist kompatibel zu iTextSharp **4.x und 5.x**. DLLs sind gitignored und muessen einmalig entsperrt werden (`Unblock-File`, sonst `0x80131515`) - Details in `Assets\lib\README.md`.
- **Excel-COM im Runspace:** `Invoke-Step3ReadSumme` liest Zelle J2 via `New-Object -ComObject Excel.Application` und gibt die COM-Objekte sauber frei (`ReleaseComObject`, `GC.Collect`). Laeuft im STA-Runspace, Ergebnis ueber `$sync.Result`.

## Datei-Layout

```
Main.ps1                          # Einstieg, XAML, Event-Handler, Settings, $global:-Mirror, Hint-Galerie, Apply-Strings-Aufruf, Resolve-DerivedPaths
Assets/hints/                     # Hinweisbilder (gitignored, .gitkeep). Pfade aus Config.Hints
Assets/lib/                       # itextsharp.dll (+ ggf. BouncyCastle) fuer Step2-PDF-Merge - gitignored, nur README.md/.gitkeep im Repo
Config/default.json               # Paths.{Msg,Task,Zahllauf}Folder + Skriptnamen, Behavior.SimulateCleanup, ScheduledTasks[], Hints{}, UI-Defaults (TempRoot/ExtractedFolder werden zur Laufzeit abgeleitet)
Config/strings.de.json            # statische UI-Texte (UTF-8) -> doc/TEXTE_BEARBEITEN.md
Changes/AktuellerStand.md         # Detailliertes Stand- und Bug-Log
doc/CLAUDE.MD                     # diese Datei
doc/TEXTE_BEARBEITEN.md           # Bearbeiter-Anleitung fuer die UI-Texte (Mapping-Tabelle)
doc/prompt.md
Functions/
  ConfigLoader.ps1                # Import-AppConfig, Save-AppConfig, Resolve-DerivedPaths (TempRoot/ExtractedFolder aus MsgFolder)
  DateHelpers.ps1                 # global:Get-IsoCalendarWeek + Get-ZahllaufContext (script-scope)
  Logging.ps1                     # global:Write-Log, global:Clear-Log
  RunspaceHelpers.ps1             # global:Start-RunspaceJob + Inner-Proxies (inkl. Confirm-Dialog)
  Strings.ps1                     # Import-AppStrings, Apply-Strings, $Script:StringMap (-> strings.de.json)
Modules/
  Step1_Mail.psm1                 # Invoke-Step1OpenOutlook, Invoke-Step1OpenPath, Invoke-Step1CopyPath, Invoke-Step1MarkDone
  Step2_Rechnungen.psm1           # Invoke-ExtractRechnungen, Invoke-Step2Run, Invoke-Step2OpenPath/OpenExcel, Invoke-Step2MergePdf/OpenMergedPdf, Invoke-Step2MarkDone
  Step3_Pruflauf.psm1             # Invoke-Step3Copy, Invoke-Step3ReadSumme (J2 via Excel-COM), Invoke-Step3OpenProsos, Invoke-Step3MarkDone
  Step4_Auszahlung.psm1           # Invoke-StartTaskScript, Invoke-Step4Run, Invoke-Step4OpenTask, Invoke-Step4OpenAbgleichFile, Invoke-Step4MarkDone
  Step4_ScheduledTasks.psm1       # Invoke-StartScheduledTask, Invoke-ScheduledTaskById (Aufgabenplanung on demand)
  Step5_Abschluss.psm1            # Invoke-Step5MailTemplate (-Summe/-ZielpfadLink), Invoke-Step5DeleteTemp, Invoke-Step5Finish
Scripts/                          # Originale - Param-Block + punktuelle Anpassungen
  RechnungenausMailinExcel.ps1    # KEIN BOM -> nur ASCII in Strings!
  viafintech_Task_Skript.ps1      # UTF-8 mit BOM
  MailanHaushalt.ps1              # UTF-8 mit BOM; Outlook-Entwurf, -Bezeichnung/-Summe/-ZielpfadLink-Param
Views/MainWindow.xaml
```