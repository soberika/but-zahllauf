# PDF-Bibliothek (Assets/lib)

Lege hier die Datei **`itextsharp.dll`** (iTextSharp 5.x) ab. Sie wird fuer die
Funktion **"PDFs zusammenfassen"** in Schritt 2 benoetigt
(`Invoke-Step2MergePdf` in `Modules/Step2_Rechnungen.psm1`).

## Warum diese Loesung
- Reines .NET, laeuft unter PowerShell 5.1 / .NET Framework 4.8.
- Keine Installation, keine Admin-Rechte, kein Internet auf dem Zielserver noetig.
- Die DLL wird zur Laufzeit per `Add-Type` aus genau diesem Ordner geladen.

## Beschaffung (auf einem Rechner MIT Internet)
1. NuGet-Paket **`iTextSharp`** (Version 5.5.13.x) herunterladen
   (z. B. von nuget.org -> "Download package").
2. Die `.nupkg` ist ein ZIP: oeffnen und aus `lib\` die **`itextsharp.dll`**
   entnehmen.
3. Diese `itextsharp.dll` hierher kopieren: `Assets\lib\itextsharp.dll`.

## Hinweise
- Fehlt die DLL, meldet Schritt 2 einen klaren Fehler und bricht sauber ab
  (es wird nichts geloescht oder ueberschrieben).
- Die DLL ist absichtlich **gitignored** (Binaerdatei). Auf dem Zielsystem
  einmalig ablegen.
