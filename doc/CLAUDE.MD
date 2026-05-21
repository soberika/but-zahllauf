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

## Aktueller Stand

- **Phase 1 (Skeleton):** fertig — WPF-GUI mit Sidebar-Navigation, Dark Theme, Logging (`Functions/Logging.ps1`), Auto-KW-Berechnung (`Functions/DateHelpers.ps1`), 5 Schritte als Pages.
- **Phase 2 (Skript-Integration):** fertig — die beiden Originalskripte werden aus der GUI im Runspace gestartet:
  - Schritt 2: `Scripts/RechnungenausMailinExcel.ps1` via `Invoke-ExtractRechnungen` (Optionen `-NurExcel`, `-CsvOrdner`).
  - Schritt 4: `Scripts/viafintech_Task_Skript.ps1` via `Invoke-StartTaskScript` (nicht-interaktiv, prueft Zielordner unter `Zahllauf\`).

## Architektur-Notizen

- **Runspaces:** Alle langlaufenden Skripte laufen ueber `Start-RunspaceJob` (`Functions/RunspaceHelpers.ps1`).
  - Shared State: `[hashtable]::Synchronized(@{ LogBox; Dispatcher; LogFile; Done; Error; Result })`.
  - Im Runspace existieren Proxies fuer `Write-Log`, `Write-Host` (Farbe -> Level), `Read-Host` (liefert `''`, damit Originalskripte nicht-interaktiv durchlaufen).
  - Abschluss-Erkennung via `DispatcherTimer` (250 ms), danach `OnComplete`-Callback im GUI-Thread.
- **Originalskripte in `Scripts/` werden nie veraendert** — nur aufgerufen.
- **Skript-Pfade** kommen aus `Config/default.json` (`Paths.ScriptRechnungen`, `Paths.ScriptTask`).
- **UI-Status** (`Txt2LastRun`, `Txt4Status`) wird im `OnComplete` gruen (Brush `Success`) oder rot (Brush `Danger`) gefaerbt.

## Datei-Layout

```
Main.ps1                          # Einstieg, XAML laden, Event-Handler
Config/default.json               # Pfade, Skriptnamen, UI-Defaults
Functions/
  DateHelpers.ps1                 # ISO-KW, Bezeichnung, Ordnername
  Logging.ps1                     # Write-Log (RichTextBox + Datei)
  ConfigLoader.ps1                # Import-AppConfig
  RunspaceHelpers.ps1             # Start-RunspaceJob
Modules/
  Step1_Mail.psm1
  Step2_Rechnungen.psm1           # Invoke-ExtractRechnungen, Invoke-Step2Run
  Step3_Pruflauf.psm1
  Step4_Auszahlung.psm1           # Invoke-StartTaskScript, Invoke-Step4Run
  Step5_Abschluss.psm1
Scripts/                          # Originale - NICHT veraendern
  RechnungenausMailinExcel.ps1
  viafintech_Task_Skript.ps1
Views/MainWindow.xaml
```