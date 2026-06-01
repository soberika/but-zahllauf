# =============================================================================
#  StepState.ps1
#  Persistiert den Erledigt-Status der 5 Schritte pro KW-Ordnername in einer
#  eigenen Config\state.json (haelt default.json sauber). state.json ist
#  Laufzeit-Status und gitignored.
# =============================================================================

function Get-StepStatePath {
    return (Join-Path $Script:AppRoot 'Config\state.json')
}

function Import-StepState {
    $path = Get-StepStatePath
    if (-not [System.IO.File]::Exists($path)) { return @{} }
    try {
        $raw = Get-Content -Path $path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) { return @{} }
        $obj = $raw | ConvertFrom-Json
        $ht  = @{}
        foreach ($p in $obj.PSObject.Properties) { $ht[$p.Name] = $p.Value }
        return $ht
    } catch {
        Write-Log "Schritt-Status konnte nicht gelesen werden: $($_.Exception.Message)" -Level Warning
        return @{}
    }
}

function Save-StepState {
    param([hashtable]$State)
    $path = Get-StepStatePath
    try {
        ($State | ConvertTo-Json -Depth 5) | Set-Content -Path $path -Encoding UTF8
    } catch {
        Write-Log "Schritt-Status konnte nicht gespeichert werden: $($_.Exception.Message)" -Level Warning
    }
}
