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