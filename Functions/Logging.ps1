# =============================================================================
#  Logging.ps1
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

    $box = $global:LogBox
    if (-not $box) { return }

    $hex = switch ($Level) {
        'Info'    { '#E0E0E0' }
        'Success' { '#4CAF50' }
        'Warning' { '#FFC107' }
        'Error'   { '#F44336' }
        'Debug'   { '#9E9E9E' }
    }

    $append = {
        $brush = [System.Windows.Media.SolidColorBrush]::new(
                    [System.Windows.Media.ColorConverter]::ConvertFromString($hex))
        $para  = [System.Windows.Documents.Paragraph]::new()
        $para.Margin = '0'
        $run   = [System.Windows.Documents.Run]::new($line)
        $run.Foreground = $brush
        $para.Inlines.Add($run)
        $box.Document.Blocks.Add($para)
        $box.ScrollToEnd()
    }.GetNewClosure()

    if ($box.Dispatcher.CheckAccess()) {
        & $append
    } else {
        [void]$box.Dispatcher.Invoke([Action]$append)
    }
}

function global:Clear-Log {
    if ($global:LogBox) { $global:LogBox.Document.Blocks.Clear() }
}
