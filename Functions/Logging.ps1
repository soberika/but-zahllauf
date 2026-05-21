# =============================================================================
#  Logging.ps1
#  Schreibt farbige Eintraege in die globale Log-RichTextBox (global:LogBox).
# =============================================================================

function global:Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info','Success','Warning','Error','Debug')]
        [string]$Level = 'Info'
    )

    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $line      = "[$timestamp] [$Level] $Message"

    if ($global:LogFile) {
        try { Add-Content -Path $global:LogFile -Value $line -Encoding UTF8 } catch { }
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
        $brush = New-Object System.Windows.Media.SolidColorBrush(
                    [System.Windows.Media.ColorConverter]::ConvertFromString($hex))
        $para  = New-Object System.Windows.Documents.Paragraph
        $para.Margin = '0'
        $run   = New-Object System.Windows.Documents.Run($text)
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
