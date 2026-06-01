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
│   ├── default.json               # Pfade, Skriptnamen, Behavior, UI-Optionen, Step-Liste
│   ├── strings.de.json            # statische UI-Texte (UTF-8) -> doc/TEXTE_BEARBEITEN.md
│   └── state.json                 # Erledigt-Status pro KW (gitignored, zur Laufzeit)
├── Views/
│   └── MainWindow.xaml            # WPF-Oberflaeche (Dark Theme)
├── Functions/
│   ├── DateHelpers.ps1            # ISO-KW, letzter Sonntag, Bezeichnungen
│   ├── Logging.ps1                # Write-Log + Status-Banner (Set-AppStatus)
│   ├── ConfigLoader.ps1           # JSON-Konfig laden + speichern
│   ├── RunspaceHelpers.ps1        # Start-RunspaceJob + Inner-Proxies (Confirm-Dialog)
│   ├── Strings.ps1                # statische UI-Texte (strings.de.json)
│   └── StepState.ps1              # Erledigt-Status pro KW (state.json)
├── Modules/
│   ├── Step1_Mail.psm1
│   ├── Step2_Rechnungen.psm1
│   ├── Step3_Pruflauf.psm1
│   ├── Step4_Auszahlung.psm1
│   ├── Step4_ScheduledTasks.psm1  # Windows-Tasks on demand anstossen + Status
│   └── Step5_Abschluss.psm1       # Mail-Vorlage + Abschluss-Cleanup
├── Scripts/
│   ├── RechnungenausMailinExcel.ps1
│   ├── viafintech_Task_Skript.ps1
│   └── MailanHaushalt.ps1         # Outlook-Entwurf fuer den Abschluss
├── Assets/
│   ├── hints/                     # Hinweisbilder (gitignored)
│   └── lib/                       # itextsharp.dll fuer PDF-Merge (gitignored, siehe README dort)
├── doc/                           # CLAUDE.MD (Regeln + Architektur), prompt.md
├── Changes/                       # AktuellerStand.md (Detail-Log)
└── Logs/
    └── zahllauf_YYYY-MM-DD.log    # Tageslog (automatisch)
```

## Features

- Dark-Theme WPF-Oberflaeche, 1320x840
- Top-Bar mit aktuellem Datum, letztem Sonntag und einer Fortschrittsanzeige
  ("x / 5 erledigt") sowie einer ComboBox zur Auswahl der bearbeiteten
  Kalenderwoche (letzte 4 KWs, Default = vergangene Woche)
- Linke Sidebar mit 5 Schritten (aktiver Schritt hervorgehoben, erledigte
  Schritte gruen markiert)
- Hauptbereich wechselt per `Visibility` zwischen den Step-Pages
- Status-Banner oberhalb des Inhalts fuer prominente Fehler-/Hinweismeldungen
  (nicht erreichbare Ordner, Runspace-Fehler, KW-Wechsel)
- Log-Fenster unten (RichTextBox), farbig nach Level
  (Info / Success / Warning / Error / Debug) plus Tageslogdatei
- Automatische Berechnung beim Start; alle Werte basieren auf der
  ausgewaehlten KW (Default = vergangene Woche):
  - ISO-Kalenderwoche (Montag = Wochenbeginn)
  - ISO-Jahr (korrekt am Jahreswechsel)
  - Letzter Sonntag der Zielwoche (= Faelligkeit)
  - Bezeichnung `Viafintech vom XX.KW YY`
  - Task-Ordnername `YY_MM_DD Fuer XX.KW`

Jeder Schritt hat zusaetzlich einen **"Schritt als erledigt markieren"**-Button:
er protokolliert den Schritt, faerbt ihn in der Sidebar gruen, erhoeht den
Fortschrittszaehler und springt zum naechsten Schritt. Der Erledigt-Status wird
pro KW in `Config/state.json` gespeichert (gitignored) und beim KW-Wechsel
wieder geladen; ueber die Einstellungen laesst er sich pro KW zuruecksetzen.

### Schritte

- **Schritt 1 — Mail:** Outlook oeffnen sowie zwei Karten/Checkboxen mit eigenen
  Ordner-Buttons (Task- und Mail-Ordner); der Mail-Ordnerpfad laesst sich
  zusaetzlich in die Zwischenablage kopieren.
- **Schritt 2 — Rechnungen:** ruft `RechnungenausMailinExcel.ps1` im Runspace,
  extrahiert `.msg`-Anhaenge, erzeugt `alleRechnungen.xlsx` und verschiebt sie
  nach `TaskFolder`. Excel und Extraktionsordner lassen sich direkt oeffnen.
  - *Bereitschaftspruefung:* zeigt beim Oeffnen, wie viele `.msg`-Dateien
    vorliegen; bei 0 Dateien oder nicht erreichbarem Ordner ist
    "Rechnungen extrahieren" gesperrt.
  - *PDFs zusammenfassen:* fasst die extrahierten Einzel-PDFs (in Mail-/
    Extraktionsreihenfolge) zu `alleRechnungen_Anhaenge.pdf` im `TaskFolder`
    zusammen. Nutzt `itextsharp.dll` aus `Assets\lib\` (reines .NET, keine
    Installation/Internet) — die DLL ist gitignored und einmalig dort abzulegen
    (siehe `Assets\lib\README.md`).
- **Schritt 3 — Prueflauf:** Bezeichnung in Zwischenablage, Prosos starten.
  - *Gesamtsumme ermitteln:* liest die Summe aus Zelle J2 der
    `alleRechnungen.xlsx` (Excel-COM) und zeigt sie formatiert in Euro an.
- **Schritt 4 — Auszahlung:** ruft `viafintech_Task_Skript.ps1`, baut den
  Task-Ordner und verschiebt ihn ins Zahllauf-Ziel. Existiert das Ziel bereits,
  fragt ein Ja/Nein-Dialog vor dem Ueberschreiben.
  - *Abgleich der Gesamtsummen:* nach erfolgreichem Lauf lassen sich die
    `alleRechnungen.xlsx` und die Haushaltsstellen-Gesamtbetraege-PDF aus dem
    Zielordner direkt zum Abgleich oeffnen.
  - *Aufgabenplanung (on demand):* startet in der Windows-Aufgabenplanung
    registrierte Tasks (z. B. die OPEN-PROSOZ-Zahllistenerstellung) per Knopf
    und meldet den Ausfuehrungsstatus zurueck. Ein Ja/Nein-Dialog bestaetigt
    jeden Start; der Status wird ausschliesslich ueber `Get-ScheduledTaskInfo`
    geprueft. **Es liegen keine Passwoerter, Argumente oder Task-XML im Repo** —
    die Aufgabe ist bereits in der Aufgabenplanung registriert und wird nur
    ueber Pfad + Name referenziert.
- **Schritt 5 — Abschluss** (in zwei getrennte Aktionen aufgeteilt):
  - *Mail-Vorlage oeffnen* erzeugt einen Outlook-Entwurf
    (`MailanHaushalt.ps1`) mit Betreff
    `<Bezeichnung> Prueflauf - steht zur weiteren Kontrolle bereit`; der Body
    enthaelt die ermittelte Gesamtsumme und einen Link auf den Zielordner.
  - *Temporaere Dateien loeschen* entfernt die `.msg`-Dateien und die
    extrahierten Anhaenge. Per **Simulation-Toggle** in den Einstellungen wird
    das Loeschen zunaechst nur protokolliert.
  - *Abschliessen* sichert das Tageslog in den Taskplaner-Ordner und
    protokolliert die veranlasste Uebergabe.

## Konfiguration

Alle Pfade und Optionen liegen in `Config/default.json` und sind ueber die
Seite **Einstellungen** in der GUI editierbar (Ordner-Browser + Speichern):

- `Paths.MsgFolder`, `Paths.TaskFolder`, `Paths.ZahllaufFolder` — Arbeitsordner
- `Paths.ScriptRechnungen`, `Paths.ScriptTask`, `Paths.ScriptMail` — Skriptnamen
- `Behavior.SimulateCleanup` — `true` = Abschluss-Cleanup nur simulieren
- `ScheduledTasks[]` — anstossbare Windows-Tasks, je `Id`, `Label`, `TaskPath`,
  `TaskName` und optional `ResultFile` (nur Existenz-/Aenderungszeit-Pruefung).
  Nur Pfad/Name/Label — keine Geheimnisse, daher Git-unbedenklich.

Aenderungen werden persistiert und beim naechsten Start wieder geladen. Die
Arbeitspfade `TempRoot` und `ExtractedFolder` werden zur Laufzeit aus
`MsgFolder` abgeleitet (Elternordner bzw. `…\extracted_attachments`) und folgen
damit automatisch einem geaenderten `MsgFolder`.

Fuer die PDF-Zusammenfassung in Schritt 2 muss einmalig `itextsharp.dll` unter
`Assets\lib\` abgelegt (und nach dem Download per `Unblock-File` entsperrt)
werden — Details in `Assets\lib\README.md`. Die DLL ist absichtlich gitignored.

## Hinweise fuer Beitragende

- **PowerShell 5.1 / WPF (.NET Framework)** — kein PowerShell 7.
- **Encoding:** BOM-lose `.ps1` (z. B. `RechnungenausMailinExcel.ps1`) nur in
  ASCII halten (Umlaute als `ue`/`ae`/`oe`, kein En-Dash) — sonst bricht der
  PS5.1-Parser. Details und weitere Architektur-Regeln in `doc/CLAUDE.MD`.

## Offen / geplant

- Tooltips/Erklaerungen fuer Fachbegriffe (Konzept steht, zurueckgestellt)
- Bereitschaftspruefung auf weitere Schritte ausweiten (aktuell nur Schritt 2)
- Validierung der Task-Ordner-Inhalte vertiefen

Bereits umgesetzt: Fortschrittsanzeige pro Schritt (Sidebar + Top-Bar),
Persistenz des Schritt-Status pro KW (`Config/state.json`), Pfad-Check beim
Start und Status-Banner fuer Fehler/Hinweise.
