# =============================================================================
#  ConfigLoader.ps1
#  Laedt und speichert die Konfiguration unter Config\default.json.
# =============================================================================

function Get-AppConfigPath {
    return (Join-Path $Script:AppRoot 'Config\default.json')
}

function Import-AppConfig {
    [CmdletBinding()]
    param(
        [string]$Path = (Get-AppConfigPath)
    )

    if (-not (Test-Path $Path)) {
        throw "Config-Datei nicht gefunden: $Path"
    }

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}

function Save-AppConfig {
    [CmdletBinding()]
    param(
        $Config = $Script:Config,
        [string]$Path = (Get-AppConfigPath)
    )

    $json = $Config | ConvertTo-Json -Depth 8
    Set-Content -Path $Path -Value $json -Encoding UTF8
}
