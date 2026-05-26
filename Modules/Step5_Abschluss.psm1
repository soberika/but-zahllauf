# =============================================================================
#  Step5_Abschluss.psm1
#  Schritt 5 - Abschluss.
# =============================================================================

function Invoke-Step5MailTemplate {
    $scriptName = $global:Config.Paths.ScriptMail
    $scriptPath = Join-Path $global:AppRoot ("Scripts\" + $scriptName)

    if (-not (Test-Path -LiteralPath $scriptPath)) {
        Write-Log "Mail-Skript nicht gefunden: $scriptPath" -Level Error
        return
    }

    $bezeichnung = (Get-ZahllaufContext).Bezeichnung
    Write-Log "Oeffne Mail-Vorlage fuer: $bezeichnung" -Level Info

    $params = @{
        ScriptPath  = $scriptPath
        Bezeichnung = $bezeichnung
    }

    Start-RunspaceJob -Parameters $params -OnComplete {
        param($sync)
        if ($sync.Error) {
            Write-Log "Mail-Vorlage: Fehler - $($sync.Error.Exception.Message)" -Level Error
        } else {
            Write-Log "Mail-Vorlage geoeffnet." -Level Success
        }
    } -ScriptBlock {
        & $ScriptPath -Bezeichnung $Bezeichnung
    } | Out-Null
}

function Invoke-Step5Finish {
    Write-Log "Zahllauf abgeschlossen." -Level Success
}

Export-ModuleMember -Function *
