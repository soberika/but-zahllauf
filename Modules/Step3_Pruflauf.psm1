# =============================================================================
#  Step3_Pruflauf.psm1
#  Schritt 3 - Pruflauf in Prosos. Platzhalter fuer Phase 1.
# =============================================================================

function Invoke-Step3Copy {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        Write-Log "Keine Bezeichnung zum Kopieren vorhanden." -Level Warning
        return
    }
    Set-Clipboard -Value $Text
    Write-Log "Bezeichnung kopiert: $Text" -Level Success
}

function Invoke-Step3OpenProsos {
    Write-Log "Prosos oeffnen (PLATZHALTER) - wird in Phase 2 integriert." -Level Info
}

Export-ModuleMember -Function *
