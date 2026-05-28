# UI-Texte bearbeiten

Diese Anleitung richtet sich an alle, die die sichtbaren Texte der Oberflaeche
aendern wollen, ohne den Programmcode zu verstehen.

## Wo liegen die Texte?

Alle statischen Texte der Oberflaeche stehen in einer einzigen Datei:

```
Config/strings.de.json
```

Das ist eine einfache Textdatei im JSON-Format. Pro Zeile steht ein
**Schluessel** (links, in Anfuehrungszeichen) und der **sichtbare Text**
(rechts, in Anfuehrungszeichen), getrennt durch einen Doppelpunkt:

```json
"Step1.Title": "Schritt 1 - Mail",
"Btn.Step1.OpenPath": "Pfad öffnen",
```

## Wie aendere ich einen Text?

1. `Config/strings.de.json` in einem Editor oeffnen, der **UTF-8** speichert
   (z. B. VS Code oder Notepad++; bei Windows-Notepad beim Speichern als
   Kodierung "UTF-8" waehlen).
2. Den gewuenschten Schluessel suchen (siehe Tabelle unten) und **nur den Text
   rechts vom Doppelpunkt** zwischen den Anfuehrungszeichen aendern.
3. Speichern.
4. **Die App neu starten** (`Main.ps1`). Texte werden nur beim Start geladen,
   nicht zur Laufzeit.

Wichtig:
- Die Anfuehrungszeichen, den Doppelpunkt und das Komma am Zeilenende **nicht**
  entfernen.
- Den **Schluessel links nicht umbenennen** (der ist mit dem Code verdrahtet).
- Umlaute (ae/oe/ue/ss) duerfen direkt geschrieben werden, solange die Datei als
  UTF-8 gespeichert ist (siehe Abschnitt "Encoding").
- Soll im Text selbst ein Anfuehrungszeichen vorkommen, muss es als `\"`
  geschrieben werden (Beispiel: `Step1.Anleitung.3`).

## Mapping-Tabelle: Schluessel -> Ort in der GUI

| Schluessel | Ort in der GUI | Beispieltext |
| --- | --- | --- |
| `Window.Title` | Fenstertitel (Titelleiste) | viafintech Zahllauf Dashboard |
| `Brand.Title` | Kopfzeile oben links, gross | viafintech Zahllauf |
| `Brand.Subtitle` | Kopfzeile oben links, klein | Dashboard v1.0 |
| `Top.KW` | Kopfzeile, Karte "Kalenderwoche" (Label) | Kalenderwoche |
| `Top.Today` | Kopfzeile, Karte "Heute" (Label) | Heute |
| `Top.LastSunday` | Kopfzeile, Karte "Letzter Sonntag" (Label) | Letzter Sonntag |
| `Top.WeekSelect` | Kopfzeile rechts, Label ueber Wochen-Auswahl | Bearbeitete KW |
| `Sidebar.Section` | Seitenleiste, Ueberschrift | ABLAUF |
| `Sidebar.Step1` | Seitenleiste, Button Schritt 1 | 1.  Mail |
| `Sidebar.Step2` | Seitenleiste, Button Schritt 2 | 2.  Rechnungen |
| `Sidebar.Step3` | Seitenleiste, Button Schritt 3 | 3.  Prüflauf |
| `Sidebar.Step4` | Seitenleiste, Button Schritt 4 | 4.  Auszahlungslauf |
| `Sidebar.Step5` | Seitenleiste, Button Schritt 5 | 5.  Abschluss |
| `Sidebar.Settings` | Seitenleiste, Button Einstellungen | Einstellungen |
| `Step1.Title` | Schritt 1, Ueberschrift (H1) | Schritt 1 - Mail |
| `Step1.Desc` | Schritt 1, Beschreibung unter Titel | Wochenmail von viafintech prüfen ... |
| `Step1.AnleitungTitle` | Schritt 1, Karten-Ueberschrift (H2) | Anleitung |
| `Step1.Anleitung.1` | Schritt 1, Anleitung Zeile 1 | 1. Outlook öffnen ... |
| `Step1.Anleitung.2` | Schritt 1, Anleitung Zeile 2 | 2. Alle relevanten .msg-Dateien ... |
| `Step1.Anleitung.3` | Schritt 1, Anleitung Zeile 3 | 3. Unten "Pfad öffnen" klicken ... |
| `Btn.Step1.OpenPath` | Schritt 1, Button | Pfad öffnen |
| `Btn.Step1.Done` | Schritt 1, Button | Schritt als erledigt markieren |
| `Step2.Title` | Schritt 2, Ueberschrift (H1) | Schritt 2 - Rechnungen extrahieren |
| `Step2.Desc` | Schritt 2, Beschreibung unter Titel | Verarbeitet die .msg-Dateien ... |
| `Step2.OptionenTitle` | Schritt 2, Karten-Ueberschrift (H2) | Optionen |
| `Chk.Step2.NurExcel` | Schritt 2, CheckBox | Nur Excel erstellen (CSV-Dateien ...) |
| `Step2.PreviewTitle` | Schritt 2, Ueberschrift (H2) | Vorschau Excel-Bezeichnung |
| `Step2.LastRunTitle` | Schritt 2, Ueberschrift (H2) | Letzter Lauf |
| `Btn.Step2.Run` | Schritt 2, Button | Rechnungen extrahieren |
| `Btn.Step2.OpenPath` | Schritt 2, Button | Ergebnis-Ordner öffnen |
| `Step3.Title` | Schritt 3, Ueberschrift (H1) | Schritt 3 - Prüflauf in Prosos |
| `Step3.Desc` | Schritt 3, Beschreibung unter Titel | Bezeichnung und Fälligkeit ... |
| `Step3.BezeichnungTitle` | Schritt 3, Ueberschrift (H2) | Bezeichnung |
| `Step3.FaelligkeitTitle` | Schritt 3, Ueberschrift (H2) | Fälligkeit (letzter Sonntag) |
| `Btn.Step3.Copy` | Schritt 3, Button | Bezeichnung in Zwischenablage |
| `Btn.Step3.Prosos` | Schritt 3, Button | Prosos öffnen |
| `Step3.ChecklistTitle` | Schritt 3, Karten-Ueberschrift (H2) | Checkliste |
| `Chk.Step3.Done0` | Schritt 3, CheckBox 1 | Prüflauf in Prosos angelegt |
| `Chk.Step3.Done1` | Schritt 3, CheckBox 2 | Bezeichnung gesetzt |
| `Chk.Step3.Done2` | Schritt 3, CheckBox 3 | Fälligkeit auf letzten Sonntag gestellt |
| `Chk.Step3.Done3` | Schritt 3, CheckBox 4 | Prüflauf erfolgreich abgeschlossen |
| `Step4.Title` | Schritt 4, Ueberschrift (H1) | Schritt 4 - Auszahlungslauf |
| `Step4.Desc` | Schritt 4, Beschreibung unter Titel | Baut den Task-Ordner ... |
| `Step4.AufgabenTitle` | Schritt 4, Karte "Aufgabenplanung" (H2) | Aufgabenplanung |
| `Step4.AufgabenDesc` | Schritt 4, Beschreibung in der Karte | Startet einen registrierten Windows-Task ... |
| `Step4.OrdnerTitle` | Schritt 4, Ueberschrift (H2) | Vorschau Ordnername |
| `Step4.StatusTitle` | Schritt 4, Ueberschrift (H2) | Task-Ordner Status |
| `Btn.Step4.Run` | Schritt 4, Button | Auszahlungslauf starten |
| `Btn.Step4.OpenTask` | Schritt 4, Button | Task-Ordner öffnen |
| `Step5.Title` | Schritt 5, Ueberschrift (H1) | Schritt 5 - Abschluss |
| `Step5.Desc` | Schritt 5, Beschreibung unter Titel | Mail an Buchhaltung versenden ... |
| `Step5.LetzteTitle` | Schritt 5, Karten-Ueberschrift (H2) | Letzte Schritte |
| `Chk.Step5.Done0` | Schritt 5, CheckBox 1 | Zahlungsdatei in Zahllauf-Ordner kopiert |
| `Chk.Step5.Done1` | Schritt 5, CheckBox 2 | Bestätigung von Buchhaltung erhalten |
| `Chk.Step5.Done2` | Schritt 5, CheckBox 3 | Backup gesichert |
| `Btn.Step5.MailTemplate` | Schritt 5, Button | Mail-Vorlage öffnen |
| `Btn.Step5.Finish` | Schritt 5, Button | Alles abschließen |
| `Settings.Title` | Einstellungen, Ueberschrift (H1) | Einstellungen - Pfade |
| `Settings.Desc` | Einstellungen, Beschreibung unter Titel | Zentrale Ordner fuer Msg-Eingang ... |
| `Settings.MsgTitle` | Einstellungen, Ueberschrift (H2) | Msg-Ordner (.msg-Dateien) |
| `Settings.TaskTitle` | Einstellungen, Ueberschrift (H2) | Task-Ordner (alleRechnungen.xlsx ...) |
| `Settings.ZahllaufTitle` | Einstellungen, Ueberschrift (H2) | Zahllauf-Ordner (Zielordner) |
| `Btn.Settings.Browse` | Einstellungen, alle drei "Ordner waehlen"-Buttons | Ordner auswählen |
| `Btn.Settings.Save` | Einstellungen, Button | Einstellungen speichern |
| `Settings.BehaviorTitle` | Einstellungen, Karten-Ueberschrift (H2) | Abschluss-Optionen |
| `Chk.Settings.Simulate` | Einstellungen, CheckBox | Cleanup simulieren (kein echtes Löschen) |
| `Settings.SimulateDesc` | Einstellungen, Hinweis unter der CheckBox | Wenn aktiv: .msg-Dateien ... |
| `Right.Section` | Rechtes Panel, Ueberschrift | GLOBALE PARAMETER |
| `Right.KW` | Rechtes Panel, Label "KW" | KW |
| `Right.Year` | Rechtes Panel, Label "Jahr" | Jahr |
| `Right.Bez` | Rechtes Panel, Label | Bezeichnung-Vorschau |
| `Right.Ord` | Rechtes Panel, Label | Ordnername-Vorschau |
| `Log.Section` | Log-Bereich unten, Ueberschrift | LOG |
| `Btn.Log.Open` | Log-Bereich, Button | Log-Datei öffnen |
| `Btn.Log.Clear` | Log-Bereich, Button | Log löschen |
| `Hints.ExpanderHeader` | Alle Schritte, Titel des Hinweis-Klappbereichs | Bebilderte Hinweise |

Hinweis: `Btn.Settings.Browse` und `Hints.ExpanderHeader` werden absichtlich
mehrfach verwendet. Wer den Text aendert, aendert ihn an allen zugehoerigen
Stellen gleichzeitig.

## Was passiert, wenn ein Schluessel fehlt?

Fehlt ein Schluessel in der JSON-Datei, bleibt der im XAML hinterlegte
Standardtext stehen und es erscheint eine Warnung im Log-Bereich
(z. B. `Strings: Schluessel fehlt -> 'Step1.Title'`). Die App stuerzt nicht ab.

## Wie fuege ich einen ganz neuen Text hinzu?

Das betrifft nur Faelle, in denen ein **neues** Textelement in der Oberflaeche
ergaenzt werden soll. Es sind drei Schritte noetig (Code-Kenntnis hilfreich):

1. **JSON:** In `Config/strings.de.json` eine neue Zeile mit Schluessel und Text
   anlegen.
2. **XAML:** Dem Element in `Views/MainWindow.xaml` ein `x:Name` geben
   (z. B. `x:Name="Txt9Neu"`).
3. **Mapping:** In `Functions/Strings.ps1` in der Tabelle `$Script:StringMap`
   eine Zeile ergaenzen:
   ```powershell
   @{ N = 'Txt9Neu'; P = 'Text'; K = 'Mein.Neuer.Schluessel' }
   ```
   - `N` = der `x:Name` aus dem XAML.
   - `P` = Ziel-Property: `Text` (TextBlock), `Content` (Button/CheckBox) oder
     `Header` (Expander).
   - `K` = der Schluessel aus der JSON-Datei.

Danach App neu starten.

## Welche Texte sind NICHT in der JSON?

Bewusst **nicht** ausgelagert sind alle Texte, die sich zur Laufzeit aendern
(sie wuerden sonst sofort wieder ueberschrieben). Diese werden weiterhin direkt
im Code gesetzt:

| Text / Element | Wo im Code |
| --- | --- |
| KW / Heute / Letzter Sonntag (Werte) | `Main.ps1`, Funktion `Update-Context` |
| Bezeichnung, Jahr, Ordnername-Vorschau (Werte, rechtes Panel) | `Main.ps1`, `Update-Context` |
| Vorschau Excel-Bezeichnung (`Txt2Preview`) | `Main.ps1`, `Update-Context` |
| Bezeichnung Schritt 3 (`Txt3Bezeichnung`) | `Main.ps1`, `Update-Context` |
| Vorschau Ordnername Schritt 4 (`Txt4Ordner`) | `Main.ps1`, `Update-Context` |
| Checklisten-Zaehler "x/4 erledigt" (`Txt3Checklist`) | `Main.ps1`, `Update-ChecklistCounter` |
| "Letzter Lauf"-Status (`Txt2LastRun`) | `Main.ps1` (Btn2Run) + `Modules/Step2_Rechnungen.psm1` (OnComplete) |
| Task-Status (`Txt4Status`, `Txt4TaskStatus`) | `Modules/Step4_*.psm1` (OnComplete) |
| Einstellungen-Status (`TxtSetStatus`) | `Main.ps1` (BtnSetSave) |
| Buttons der Aufgabenplanung (`Sp4Tasks`) | aus `Config/default.json` -> `ScheduledTasks[].Label` |
| Bildunterschriften der Hinweis-Galerie | aus `Config/default.json` -> `Hints` |

Diese Texte gehoeren also entweder in den Code (dynamische Statusmeldungen) oder
in `Config/default.json` (Task-Labels, Hinweisbild-Beschriftungen).

## Encoding (wichtig)

`Config/strings.de.json` muss als **UTF-8** gespeichert werden. Dann sind echte
Umlaute (ae/oe/ue/ss als ä/ö/ü/ß) erlaubt und werden korrekt angezeigt. Wird die
Datei versehentlich in einer anderen Kodierung gespeichert, koennen Umlaute
falsch dargestellt werden. Im Zweifel im Editor die Kodierung pruefen und erneut
als UTF-8 speichern.

> Hinweis fuer Entwickler: Diese UTF-8-Freiheit gilt nur fuer JSON- und
> XAML-Dateien. Die Encoding-Regeln fuer BOM-lose `.ps1`-Dateien (nur ASCII in
> Strings) aus `doc/CLAUDE.MD` bleiben davon unberuehrt.

## Nach jeder Aenderung

Die App (`Main.ps1`) **neu starten** - die Texte werden nur einmal beim Start
aus `Config/strings.de.json` geladen.
