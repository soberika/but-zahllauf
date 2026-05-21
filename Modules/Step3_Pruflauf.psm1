# =============================================================================
#  Step3_Pruflauf.psm1
#  Schritt 3 - Pruflauf in Prosos.
# =============================================================================

function Invoke-Step3Copy {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        Write-Log "Keine Bezeichnung zum Kopieren vorhanden." -Level Warning
        return
    }

    try {
        [System.Windows.Clipboard]::SetText($Text)
        Write-Log "Bezeichnung kopiert: $Text" -Level Success
    } catch {
        Write-Log "Clipboard-Fehler: $($_.Exception.Message)" -Level Error
    }
}

function Invoke-Step3OpenProsos {
    $exe  = 'C:\Program Files (x86)\PROSOZ Herten\OPEN PROSOZ\Anwendungen\OpenStarter.exe'
    $args = '/app OpenClient.exe'

    if (Test-Path -LiteralPath $exe) {
        Start-Process -FilePath $exe -ArgumentList $args
        Write-Log "Prosos gestartet." -Level Info
    } else {
        Write-Log "Prosos-Exe nicht gefunden: $exe" -Level Warning
    }
}

Export-ModuleMember -Function *
