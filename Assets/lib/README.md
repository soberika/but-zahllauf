# PDF-Bibliothek (Assets/lib)

Lege hier die Datei **`itextsharp.dll`** (iTextSharp 5.x) ab. Sie wird fuer die
Funktion **"PDFs zusammenfassen"** in Schritt 2 benoetigt
(`Invoke-Step2MergePdf` in `Modules/Step2_Rechnungen.psm1`).

## Warum diese Loesung
- Reines .NET, laeuft unter PowerShell 5.1 / .NET Framework 4.8.
- Keine Installation, keine Admin-Rechte, kein Internet auf dem Zielserver noetig.
- Die DLL wird zur Laufzeit per `Add-Type` aus genau diesem Ordner geladen.

## Beschaffung (auf einem Rechner MIT Internet)

### Variante A (empfohlen): iTextSharp 4.1.6 - EINE DLL, keine Abhaengigkeit
1. NuGet-Paket **`iTextSharp.LGPLv2.Core`** ist .NET-Core; stattdessen die
   klassische **`iTextSharp`-Version 4.1.6.0** besorgen (letzte LGPL-Version,
   z. B. ueber das NuGet-Paket `iTextSharp` 4.1.6 oder ein Archiv).
2. Die enthaltene **`itextsharp.dll`** hierher kopieren: `Assets\lib\itextsharp.dll`.
3. Fertig - 4.1.6 ist eine einzige, in sich geschlossene DLL ohne weitere
   Abhaengigkeiten.

### Variante B: iTextSharp 5.5.13 - braucht zusaetzlich BouncyCastle
1. NuGet-Paket **`iTextSharp`** (5.5.13.x) herunterladen, `.nupkg` als ZIP
   oeffnen, aus `lib\` die **`itextsharp.dll`** entnehmen -> `Assets\lib\`.
2. Zusaetzlich **`BouncyCastle.Crypto.dll`** besorgen (NuGet-Paket
   `BouncyCastle` 1.8.x) und ebenfalls nach `Assets\lib\` legen.
   Ohne sie schlaegt das Laden mit "Mindestens ein Typ ... kann nicht geladen
   werden" fehl. Der Code laedt alle DLLs aus diesem Ordner automatisch mit.

### WICHTIG fuer beide Varianten: DLLs entsperren
Windows blockiert aus dem Internet geladene Dateien. Pro DLL einmalig:
```powershell
Get-ChildItem ".\Assets\lib\*.dll" | Unblock-File
```
Alternativ: Datei im Explorer -> Rechtsklick -> Eigenschaften -> unten
"Zulassen" anhaeken -> OK. Ohne diesen Schritt: Fehler 0x80131515.

## Hinweise
- Fehlt die DLL, meldet Schritt 2 einen klaren Fehler und bricht sauber ab
  (es wird nichts geloescht oder ueberschrieben).
- Die DLL ist absichtlich **gitignored** (Binaerdatei). Auf dem Zielsystem
  einmalig ablegen.
