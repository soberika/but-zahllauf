# =============================================================================
#  Step2_Rechnungen.psm1
#  Schritt 2 - alleRechnungen.ps1. Platzhalter fuer Phase 1.
# =============================================================================

function Invoke-Step2Run {
    param(
        [switch]$NurExcel,
        [switch]$NurExcelCsv
    )

    Write-Log "Schritt 2: alleRechnungen.ps1 (PLATZHALTER) - NurExcel=$NurExcel NurExcelCsv=$NurExcelCsv" -Level Info
    Write-Log "Wird in Phase 2 mit Runspace + Script-Aufruf integriert." -Level Debug
}

function Invoke-Step2OpenPath {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Start-Process explorer.exe -ArgumentList $Path
        Write-Log "Ergebnis-Ordner geoeffnet: $Path" -Level Info
    } else {
        Write-Log "Pfad nicht erreichbar: $Path" -Level Warning
    }
}

Export-ModuleMember -Function *
