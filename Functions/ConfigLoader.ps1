# =============================================================================
#  ConfigLoader.ps1
#  Laedt die Konfiguration aus Config\default.json und liefert ein PSCustomObject.
# =============================================================================

function Import-AppConfig {
    [CmdletBinding()]
    param(
        [string]$Path = (Join-Path $Script:AppRoot 'Config\default.json')
    )

    if (-not (Test-Path $Path)) {
        throw "Config-Datei nicht gefunden: $Path"
    }

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}
