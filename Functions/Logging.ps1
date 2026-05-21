# =============================================================================
#  Logging.ps1
#  Pure-.NET-Implementierung, robust gegen kaputten PS-Cmdlet-Lookup
#  (z.B. nach Runspace-Disposal in WPF-Event-Handlern).
# =============================================================================

function global:Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info','Success','Warning','Error','Debug')]
        [string]$Level = 'Info'
    )

    $timestamp = [DateTime]::Now.ToString('HH:mm:ss')
    $line      = "[$timestamp] [$Level] $Message"

    if ($global:LogFile) {
        try {
            [System.IO.File]::AppendAllText(
                $global:LogFile,
                ($line + "`r`n"),
                [System.Text.Encoding]::UTF8)
        } catch { }
    }

    if (-not $global:LogBox) { return }

    $color = switch ($Level) {
        'Info'    { '#E0E0E0' }
        'Success' { '#4CAF50' }
        'Warning' { '#FFC107' }
        'Error'   { '#F44336' }
        'Debug'   { '#9E9E9E' }
    }

    $action = {
        param($box, $text, $hex)
        $brush = [System.Windows.Media.SolidColorBrush]::new(
                    [System.Windows.Media.ColorConverter]::ConvertFromString($hex))
        $para  = [System.Windows.Documents.Paragraph]::new()
        $para.Margin = '0'
        $run   = [System.Windows.Documents.Run]::new($text)
        $run.Foreground = $brush
        $para.Inlines.Add($run)
        $box.Document.Blocks.Add($para)
        $box.ScrollToEnd()
    }

    if ($global:LogBox.Dispatcher.CheckAccess()) {
        & $action $global:LogBox $line $color
    } else {
        $global:LogBox.Dispatcher.Invoke($action, @($global:LogBox, $line, $color))
    }
}

function global:Clear-Log {
    if ($global:LogBox) { $global:LogBox.Document.Blocks.Clear() }
}
