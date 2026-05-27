# viafintech Zahllauf Dashboard

Modulare WPF-GUI fuer den monatlichen viafintech-Zahllaufprozess. Fuehrt den
Maintainer in 5 Schritten durch den Ablauf; langlaufende Originalskripte
laufen in eigenen STA-Runspaces, damit die Oberflaeche reaktiv bleibt.

Alle 5 Schritte sind integriert: Mail-Hinweise, Rechnungs-Extraktion
(`RechnungenausMailinExcel.ps1`), Prueflauf-Vorbereitung, Auszahlungslauf
(`viafintech_Task_Skript.ps1`) und Abschluss (Outlook-Entwurf + Cleanup).

## Schnellstart

Voraussetzungen:

- Windows 10/11
- PowerShell **5.1 Desktop** (kein PowerShell 7 — WPF braucht .NET Framework)
- Keine Admin-Rechte noetig

Start:

```powershell
cd <projektordner>
.\Main.ps1
```

Das Skript startet sich bei Bedarf automatisch im STA-Modus neu.

## Ordnerstruktur

```
.
├── Main.ps1                       # Einstiegspunkt — UI laden + Event-Handler + Settings
├── README.md
├── Config/
│   └── default.json               # Pfade, Skriptnamen, Behavior, UI-Optionen, Step-Liste
├── Views/
│   └── MainWindow.xaml            # WPF-Oberflaeche (Dark Theme)
├── Functions/
│   ├── DateHelpers.ps1            # ISO-KW, letzter Sonntag, Bezeichnungen
│   ├── Logging.ps1                # Write-Log + farbige Ausgabe in RichTextBox
│   ├── ConfigLoader.ps1           # JSON-Konfig laden + speichern
│   └── RunspaceHelpers.ps1        # Start-RunspaceJob + Inner-Proxies (Confirm-Dialog)
├── Modules/
│   ├── Step1_Mail.psm1
│   ├── Step2_Rechnungen.psm1
│   ├── Step3_Pruflauf.psm1
│   ├── Step4_Auszahlung.psm1
│   └── Step5_Abschluss.psm1       # Mail-Vorlage + Abschluss-Cleanup
├── Scripts/
│   ├── RechnungenausMailinExcel.ps1
│   ├── viafintech_Task_Skript.ps1
│   └── MailanHaushalt.ps1         # Outlook-Entwurf fuer den Abschluss
├── doc/                           # CLAUDE.MD (Regeln + Architektur), prompt.md
├── Changes/                       # AktuellerStand.md (Detail-Log)
└── Logs/
    └── zahllauf_YYYY-MM-DD.log    # Tageslog (automatisch)
```

## Features

- Dark-Theme WPF-Oberflaeche, 1320x840
- Top-Bar mit aktuellem Datum und letztem Sonntag sowie einer ComboBox
  zur Auswahl der bearbeiteten Kalenderwoche (letzte 4 KWs, Default = vergangene Woche)
- Linke Sidebar mit 5 Schritten (aktiver Schritt visuell hervorgehoben)
- Hauptbereich wechselt per `Visibility` zwischen den Step-Pages
- Rechtes Panel mit globalen Parametern (KW, Jahr, Bezeichnungs-Vorschau,
  Ordnernamens-Vorschau)
- Log-Fenster unten (RichTextBox), farbig nach Level
  (Info / Success / Warning / Error / Debug) plus Tageslogdatei
- Automatische Berechnung beim Start; alle Werte basieren auf der
  ausgewaehlten KW (Default = vergangene Woche):
  - ISO-Kalenderwoche (Montag = Wochenbeginn)
  - ISO-Jahr (korrekt am Jahreswechsel)
  - Letzter Sonntag der Zielwoche (= Faelligkeit)
  - Bezeichnung `Viafintech vom XX.KW YY`
  - Task-Ordnername `YY_MM_DD Fuer XX.KW`

### Schritte

- **Schritt 2 — Rechnungen:** ruft `RechnungenausMailinExcel.ps1` im Runspace,
  extrahiert `.msg`-Anhaenge, erzeugt `alleRechnungen.xlsx` und verschiebt sie
  nach `TaskFolder`.
- **Schritt 3 — Prueflauf:** Bezeichnung in Zwischenablage, Prosos starten.
- **Schritt 4 — Auszahlung:** ruft `viafintech_Task_Skript.ps1`, baut den
  Task-Ordner und verschiebt ihn ins Zahllauf-Ziel. Existiert das Ziel bereits,
  fragt ein Ja/Nein-Dialog vor dem Ueberschreiben.
- **Schritt 5 — Abschluss:**
  - *Mail-Vorlage oeffnen* erzeugt einen Outlook-Entwurf
    (`MailanHaushalt.ps1`) mit Betreff
    `<Bezeichnung> Prueflauf - steht zur weiteren Kontrolle bereit`.
  - *Alles abschliessen* loescht `.msg`-Dateien + `extracted_Attachements` und
    sichert das Tageslog in den Taskplaner-Ordner. Per **Simulation-Toggle** in
    den Einstellungen wird das Loeschen zunaechst nur protokolliert.

## Konfiguration

Alle Pfade und Optionen liegen in `Config/default.json` und sind ueber die
Seite **Einstellungen** in der GUI editierbar (Ordner-Browser + Speichern):

- `Paths.MsgFolder`, `Paths.TaskFolder`, `Paths.ZahllaufFolder` — Arbeitsordner
- `Paths.ScriptRechnungen`, `Paths.ScriptTask`, `Paths.ScriptMail` — Skriptnamen
- `Behavior.SimulateCleanup` — `true` = Abschluss-Cleanup nur simulieren

Aenderungen werden persistiert und beim naechsten Start wieder geladen.

## Hinweise fuer Beitragende

- **PowerShell 5.1 / WPF (.NET Framework)** — kein PowerShell 7.
- **Encoding:** BOM-lose `.ps1` (z. B. `RechnungenausMailinExcel.ps1`) nur in
  ASCII halten (Umlaute als `ue`/`ae`/`oe`, kein En-Dash) — sonst bricht der
  PS5.1-Parser. Details und weitere Architektur-Regeln in `doc/CLAUDE.MD`.

## Offen / geplant

- Fortschrittsanzeigen pro Schritt
- Persistenz des Schritt-Status zwischen Sitzungen
- Optionale Validierung der Task-Ordner-Inhalte
