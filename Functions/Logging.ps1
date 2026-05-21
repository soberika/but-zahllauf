# =============================================================================
#  Logging.ps1
#  Schreibt farbige Eintraege in die globale Log-RichTextBox (Script:LogBox).
#  Schreibt zusaetzlich in eine Tages-Logdatei unter ./Logs.
# =============================================================================

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet('Info','Success','Warning','Error','Debug')]
        [string]$Level = 'Info'
    )

    $timestamp = (Get-Date).ToString('HH:mm:ss')
    $line      = "[$timestamp] [$Level] $Message"

    # --- in Datei schreiben ---------------------------------------------------
    if ($Script:LogFile) {
        try { Add-Content -Path $Script:LogFile -Value $line -Encoding UTF8 } catch { }
    }

    # --- in WPF-RichTextBox schreiben ----------------------------------------
    if (-not $Script:LogBox) { return }

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

    if ($Script:LogBox.Dispatcher.CheckAccess()) {
        & $action $Script:LogBox $line $color
    } else {
        $Script:LogBox.Dispatcher.Invoke($action, @($Script:LogBox, $line, $color))
    }
}

function Clear-Log {
    if ($Script:LogBox) { $Script:LogBox.Document.Blocks.Clear() }
}
