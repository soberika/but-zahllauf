# =============================================================================
#  Step1_Mail.psm1
#  Schritt 1 - Mail. Platzhalter fuer Phase 1.
# =============================================================================

function Invoke-Step1OpenPath {
    param([string]$Path)

    Write-Log "Schritt 1: Oeffne Pfad '$Path'" -Level Info
    if ([System.IO.Directory]::Exists($Path)) {
        [System.Diagnostics.Process]::Start('explorer.exe', $Path) | Out-Null
    } else {
        Write-Log "Pfad nicht erreichbar: $Path" -Level Warning
    }
}

function Invoke-Step1MarkDone {
    Write-Log "Schritt 1 wurde als erledigt markiert." -Level Success
}

Export-ModuleMember -Function *
