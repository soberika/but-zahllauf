# =============================================================================
#  Step2_Rechnungen.psm1
#  Schritt 2 - ruft das Rechnungs-Skript im Runspace auf.
# =============================================================================

function Invoke-ExtractRechnungen {
    [CmdletBinding()]
    param(
        [switch]$NurExcel,
        [string]$CsvOrdner   = '',
        [string]$MsgFolder   = '',
        [string]$TaskFolder  = '',
        $StatusTextBlock,
        $BrushSuccess,
        $BrushDanger
    )

    $scriptName = $global:Config.Paths.ScriptRechnungen
    $scriptPath = Join-Path $global:AppRoot ("Scripts\" + $scriptName)

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        Write-Log "Rechnungs-Skript nicht gefunden: $scriptPath" -Level Error
        if ($StatusTextBlock) {
            $StatusTextBlock.Text = "Fehler: Skript nicht gefunden"
            if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
        }
        return
    }

    Write-Log "Starte Rechnungs-Skript: $scriptPath" -Level Info
    Write-Log ("Parameter: NurExcel={0} CsvOrdner='{1}' MsgFolder='{2}' TaskFolder='{3}'" -f `
        [bool]$NurExcel, $CsvOrdner, $MsgFolder, $TaskFolder) -Level Debug

    $params = @{
        ScriptPath = $scriptPath
        NurExcel   = [bool]$NurExcel
        CsvOrdner  = $CsvOrdner
        MsgFolder  = $MsgFolder
        TaskFolder = $TaskFolder
    }

    $onComplete = {
        param($sync)
        if ($sync.Error) {
            Write-Log "Schritt 2 abgebrochen." -Level Error
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fehler: $($sync.Error.Exception.Message)"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
            return
        }

        $xlsx = Join-Path $TaskFolder 'alleRechnungen.xlsx'
        if (Test-Path -LiteralPath $xlsx) {
            Write-Log "Schritt 2 fertig. Excel: $xlsx" -Level Success
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fertig: $xlsx"
                if ($BrushSuccess) { $StatusTextBlock.Foreground = $BrushSuccess }
            }
        } else {
            Write-Log "Schritt 2 beendet, aber alleRechnungen.xlsx nicht gefunden unter: $xlsx" -Level Warning
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Beendet, XLSX nicht gefunden"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
        }
    }.GetNewClosure()

    Start-RunspaceJob -Parameters $params -OnComplete $onComplete -ScriptBlock {
        Write-Log "Extrahiere Anhaenge / erstelle Excel..." -Level Info

        $invokeArgs = @{}
        if ($NurExcel)                                       { $invokeArgs.NurExcel     = $true }
        if (-not [string]::IsNullOrWhiteSpace($CsvOrdner))   { $invokeArgs.CsvOrdner    = $CsvOrdner }
        if (-not [string]::IsNullOrWhiteSpace($MsgFolder))   { $invokeArgs.MsgFolder    = $MsgFolder }
        if (-not [string]::IsNullOrWhiteSpace($TaskFolder))  { $invokeArgs.TargetFolder = $TaskFolder }

        & $ScriptPath @invokeArgs
        Write-Log "Rechnungs-Skript durchgelaufen." -Level Info
    } | Out-Null
}

function Invoke-Step2Run {
    [CmdletBinding()]
    param(
        [switch]$NurExcel,
        [string]$CsvOrdner = '',
        $StatusTextBlock,
        $BrushSuccess,
        $BrushDanger
    )

    # Ausgabeordner fuer die extrahierten .pdf/.csv festlegen, sonst legt das
    # Skript sie neben sich selbst ab (PSScriptRoot\extracted_attachments).
    if ([string]::IsNullOrWhiteSpace($CsvOrdner)) {
        $CsvOrdner = [string]$global:Config.Paths.ExtractedFolder
        if ([string]::IsNullOrWhiteSpace($CsvOrdner)) {
            $CsvOrdner = Join-Path $global:Config.Paths.TempRoot 'extracted_attachments'
        }
    }

    Invoke-ExtractRechnungen `
        -NurExcel:$NurExcel `
        -CsvOrdner  $CsvOrdner `
        -MsgFolder  $global:Config.Paths.MsgFolder `
        -TaskFolder $global:Config.Paths.TaskFolder `
        -StatusTextBlock $StatusTextBlock `
        -BrushSuccess $BrushSuccess `
        -BrushDanger $BrushDanger
}

function Invoke-Step2MarkDone {
    Write-Log "Schritt 2 wurde als erledigt markiert." -Level Success
}

function Invoke-Step2OpenPath {
    param([string]$Path)

    if ([System.IO.Directory]::Exists($Path)) {
        [System.Diagnostics.Process]::Start('explorer.exe', $Path) | Out-Null
        Write-Log "Ordner mit extrahierten Dateien geoeffnet: $Path" -Level Info
    } else {
        Write-Log "Pfad nicht erreichbar: $Path" -Level Warning
    }
}

function Invoke-Step2OpenExcel {
    param([string]$Path)

    if ([System.IO.File]::Exists($Path)) {
        [System.Diagnostics.Process]::Start($Path) | Out-Null
        Write-Log "Excel geoeffnet: $Path" -Level Info
    } else {
        Write-Log "Excel nicht gefunden: $Path" -Level Warning
    }
}

function Invoke-Step2OpenMergedPdf {
    param([string]$Path)

    if ([System.IO.File]::Exists($Path)) {
        [System.Diagnostics.Process]::Start($Path) | Out-Null
        Write-Log "Zusammengefasste PDF geoeffnet: $Path" -Level Info
    } else {
        Write-Log "Zusammengefasste PDF nicht gefunden: $Path" -Level Warning
    }
}

# Fasst die extrahierten Einzel-PDFs zu einer Sammeldatei zusammen.
# Reihenfolge = Mail-/Extraktionsreihenfolge (nach Erstellzeit sortiert).
# Nutzt itextsharp.dll aus Assets\lib (reines .NET, kein Install/Internet noetig).
function Invoke-Step2MergePdf {
    [CmdletBinding()]
    param(
        $StatusTextBlock,
        $BrushSuccess,
        $BrushDanger
    )

    $dllPath = Join-Path $global:AppRoot 'Assets\lib\itextsharp.dll'

    $pdfFolder = [string]$global:Config.Paths.ExtractedFolder
    if ([string]::IsNullOrWhiteSpace($pdfFolder)) {
        $pdfFolder = Join-Path $global:Config.Paths.TempRoot 'extracted_attachments'
    }
    $outPath = Join-Path $global:Config.Paths.TaskFolder 'alleRechnungen_Anhaenge.pdf'

    if (-not (Test-Path -LiteralPath $dllPath)) {
        Write-Log "PDF-Bibliothek fehlt: $dllPath - bitte itextsharp.dll einmalig unter Assets\lib\ ablegen (siehe Assets\lib\README.md)." -Level Error
        if ($StatusTextBlock) {
            $StatusTextBlock.Text = "Fehler: itextsharp.dll fehlt (Assets\lib\)"
            if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
        }
        return
    }

    Write-Log "Fasse extrahierte PDFs zusammen aus: $pdfFolder" -Level Info

    $params = @{
        DllPath   = $dllPath
        PdfFolder = $pdfFolder
        OutPath   = $outPath
    }

    $onComplete = {
        param($sync)
        if ($sync.Error) {
            Write-Log "PDF-Zusammenfassung abgebrochen." -Level Error
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fehler: $($sync.Error.Exception.Message)"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
            return
        }
        if (Test-Path -LiteralPath $outPath) {
            Write-Log "PDF-Zusammenfassung fertig: $outPath" -Level Success
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Fertig: $([System.IO.Path]::GetFileName($outPath))"
                if ($BrushSuccess) { $StatusTextBlock.Foreground = $BrushSuccess }
            }
        } else {
            Write-Log "PDF-Zusammenfassung beendet, aber keine Ausgabedatei: $outPath" -Level Warning
            if ($StatusTextBlock) {
                $StatusTextBlock.Text = "Keine PDF erzeugt (keine Anhaenge?)"
                if ($BrushDanger) { $StatusTextBlock.Foreground = $BrushDanger }
            }
        }
    }.GetNewClosure()

    Start-RunspaceJob -Parameters $params -OnComplete $onComplete -ScriptBlock {
        if (-not [System.IO.Directory]::Exists($PdfFolder)) {
            Write-Log "Ordner mit Anhaengen nicht erreichbar: $PdfFolder" -Level Warning
            return
        }

        # Reihenfolge der Mails = Extraktionsreihenfolge -> nach Erstellzeit sortieren
        $pdfs = Get-ChildItem -LiteralPath $PdfFolder -Filter '*.pdf' -File | Sort-Object CreationTime
        if (-not $pdfs -or @($pdfs).Count -eq 0) {
            Write-Log "Keine PDF-Anhaenge gefunden in: $PdfFolder" -Level Warning
            return
        }
        Write-Log "$(@($pdfs).Count) PDF-Datei(en) werden zusammengefasst." -Level Info

        if (-not ([System.Management.Automation.PSTypeName]'iTextSharp.text.Document').Type) {
            # Abhaengigkeiten (z.B. BouncyCastle.Crypto.dll bei iTextSharp 5.x)
            # zuerst laden, danach itextsharp.dll.
            $libDir = [System.IO.Path]::GetDirectoryName($DllPath)
            Get-ChildItem -LiteralPath $libDir -Filter '*.dll' -File |
                Where-Object { $_.Name -ne 'itextsharp.dll' } |
                ForEach-Object { try { Add-Type -Path $_.FullName -ErrorAction Stop } catch { } }
            try {
                Add-Type -Path $DllPath -ErrorAction Stop
            } catch {
                # Exakte fehlende Abhaengigkeit aus der ReflectionTypeLoadException melden.
                $ex = $_.Exception; $loader = $null
                while ($ex) {
                    if ($ex -is [System.Reflection.ReflectionTypeLoadException]) { $loader = $ex; break }
                    $ex = $ex.InnerException
                }
                if ($loader) {
                    $det = ($loader.LoaderExceptions | ForEach-Object { $_.Message } | Select-Object -Unique) -join '; '
                    Write-Log "PDF-Bibliothek unvollstaendig - fehlende Abhaengigkeit: $det" -Level Error
                    Write-Log "Bei iTextSharp 5.x zusaetzlich BouncyCastle.Crypto.dll nach Assets\lib\ legen (oder iTextSharp 4.1.6 verwenden, eine DLL ohne Abhaengigkeit). Siehe Assets\lib\README.md." -Level Error
                }
                throw
            }
        }

        $outDir = [System.IO.Path]::GetDirectoryName($OutPath)
        if (-not [System.IO.Directory]::Exists($outDir)) {
            [System.IO.Directory]::CreateDirectory($outDir) | Out-Null
        }

        # Seitenweises Kopieren - kompatibel zu iTextSharp 4.x UND 5.x.
        $stream = New-Object System.IO.FileStream($OutPath, [System.IO.FileMode]::Create)
        $doc    = New-Object iTextSharp.text.Document
        $copy   = New-Object iTextSharp.text.pdf.PdfCopy($doc, $stream)
        $doc.Open()
        $added = 0
        try {
            foreach ($pdf in $pdfs) {
                try {
                    $reader = New-Object iTextSharp.text.pdf.PdfReader($pdf.FullName)
                    $pages  = $reader.NumberOfPages
                    Write-Log "  + $($pdf.Name) ($pages Seite(n))" -Level Debug
                    for ($i = 1; $i -le $pages; $i++) {
                        $copy.AddPage($copy.GetImportedPage($reader, $i))
                        $added++
                    }
                    $copy.FreeReader($reader)
                    $reader.Close()
                } catch {
                    # Echten Fehler melden statt ihn vom Close() verdecken zu lassen.
                    Write-Log "  ! '$($pdf.Name)' uebersprungen: $($_.Exception.Message)" -Level Warning
                }
            }
        } finally {
            try { $doc.Close() } catch { }   # 'no pages' nicht durchreichen
            $stream.Close()
        }

        if ($added -eq 0) {
            if ([System.IO.File]::Exists($OutPath)) { [System.IO.File]::Delete($OutPath) }
            Write-Log "Keine Seiten uebernommen - PDFs evtl. beschaedigt, verschluesselt oder kein echtes PDF." -Level Warning
        } else {
            Write-Log "Zusammengefasste PDF erstellt ($added Seite(n)): $OutPath" -Level Info
        }
    } | Out-Null
}

Export-ModuleMember -Function *
