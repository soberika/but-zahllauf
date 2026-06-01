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

# Leitet die abgeleiteten Pfade aus den editierbaren Einstellungen ab, damit
# sie immer konsistent bleiben (auch nach Aenderung von MsgFolder):
#   TempRoot        = Elternordner von MsgFolder (...\_Temp)
#   ExtractedFolder = TempRoot\extracted_attachments
# Aufruf nach Import-AppConfig und nach jedem Speichern der Einstellungen.
function Resolve-DerivedPaths {
    [CmdletBinding()]
    param($Config = $Script:Config)

    $msg = [string]$Config.Paths.MsgFolder
    if ([string]::IsNullOrWhiteSpace($msg)) { return }

    $tempRoot  = Split-Path -Parent $msg
    $extracted = Join-Path $tempRoot 'extracted_attachments'

    $Config.Paths | Add-Member -NotePropertyName 'TempRoot'        -NotePropertyValue $tempRoot  -Force
    $Config.Paths | Add-Member -NotePropertyName 'ExtractedFolder' -NotePropertyValue $extracted -Force
}
