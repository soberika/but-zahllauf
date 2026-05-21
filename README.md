# viafintech Zahllauf Dashboard

Modulare WPF-GUI fuer den wochentlichen viafintech-Zahllaufprozess.
Dieses Repo enthaelt **Phase 1 — das Grundgeruest** (UI, Navigation,
Auto-Berechnungen, Logging, Platzhalter-Buttons). Die echte Integration der
bestehenden Skripte `alleRechnungen.ps1` und `viafintech_Task_Skript.ps1`
folgt in Phase 2.

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
├── Main.ps1                       # Einstiegspunkt — UI laden + Event-Handler
├── README.md
├── Config/
│   └── default.json               # Pfade, UI-Optionen, Step-Liste
├── Views/
│   └── MainWindow.xaml            # WPF-Oberflaeche (Dark Theme)
├── Functions/
│   ├── DateHelpers.ps1            # ISO-KW, letzter Sonntag, Bezeichnungen
│   ├── Logging.ps1                # Write-Log + farbige Ausgabe in RichTextBox
│   └── ConfigLoader.ps1           # JSON-Konfig laden
├── Modules/
│   ├── Step1_Mail.psm1
│   ├── Step2_Rechnungen.psm1
│   ├── Step3_Pruflauf.psm1
│   ├── Step4_Auszahlung.psm1
│   └── Step5_Abschluss.psm1
└── Logs/
    └── zahllauf_YYYY-MM-DD.log    # Tageslog (automatisch)
```

## Features (Phase 1)

- Dark-Theme WPF-Oberflaeche, 1320x840
- Top-Bar mit aktueller KW, Datum, letztem Sonntag, Refresh-Button
- Linke Sidebar mit 5 Schritten (aktiver Schritt visuell hervorgehoben)
- Hauptbereich wechselt per `Visibility` zwischen den Step-Pages
- Rechtes Panel mit globalen Parametern (KW, Jahr, Bezeichnungs-Vorschau,
  Ordnernamens-Vorschau)
- Log-Fenster unten (RichTextBox), farbig nach Level
  (Info / Success / Warning / Error / Debug) plus Tageslogdatei
- Automatische Berechnung beim Start und per Button:
  - ISO-Kalenderwoche (Montag = Wochenbeginn)
  - ISO-Jahr (korrekt am Jahreswechsel)
  - Letzter Sonntag
  - Bezeichnung `Viafintech vom XX.KW YY`
  - Task-Ordnername `YY_MM_DD Fuer XX.KW`

## Konfiguration

Alle Pfade und Optionen liegen in `Config/default.json`. Anpassen, fertig —
kein Code aendern noetig.

## Phase 2 (geplant)

- Echter Aufruf von `alleRechnungen.ps1` und `viafintech_Task_Skript.ps1`
  in einem Runspace (UI bleibt responsiv)
- Fortschrittsanzeigen pro Schritt
- Persistenz des Schritt-Status zwischen Sitzungen
- Outlook-Vorlage in Schritt 5
- Optionale Validierung der Task-Ordner-Inhalte
