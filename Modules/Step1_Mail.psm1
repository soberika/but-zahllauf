# =============================================================================
#  Step1_Mail.psm1
#  Schritt 1 - Mail. Platzhalter fuer Phase 1.
# =============================================================================

function Invoke-Step1OpenOutlook {
    Write-Log "Schritt 1: Oeffne Outlook." -Level Info
    try {
        [System.Diagnostics.Process]::Start('outlook.exe') | Out-Null
    } catch {
        Write-Log "Schritt 1: Outlook konnte nicht gestartet werden: $($_.Exception.Message)" -Level Error
    }
}

function Invoke-Step1OpenPath {
    param([string]$Path)

    Write-Log "Schritt 1: Oeffne Pfad '$Path'" -Level Info
    if ([System.IO.Directory]::Exists($Path)) {
        [System.Diagnostics.Process]::Start('explorer.exe', $Path) | Out-Null
    } else {
        Write-Log "Pfad nicht erreichbar: $Path" -Level Warning
    }
}

function Invoke-Step1CopyPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        Write-Log "Schritt 1: Kein Pfad zum Kopieren konfiguriert." -Level Warning
        return
    }
    try {
        [System.Windows.Clipboard]::SetText($Path)
        Write-Log "Schritt 1: Pfad in Zwischenablage: $Path" -Level Success
    } catch {
        Write-Log "Schritt 1: Clipboard-Fehler: $($_.Exception.Message)" -Level Error
    }
}

function Invoke-Step1MarkDone {
    Write-Log "Schritt 1 wurde als erledigt markiert." -Level Success
}

Export-ModuleMember -Function *
