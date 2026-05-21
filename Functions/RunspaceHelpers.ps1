# =============================================================================
#  RunspaceHelpers.ps1
#  Fuehrt langlaufende Skripte in einem eigenen STA-Runspace aus.
# =============================================================================

function global:Start-RunspaceJob {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [hashtable]$Parameters = @{},

        [scriptblock]$OnComplete
    )

    if (-not $global:Window) {
        throw "Start-RunspaceJob: global:Window fehlt - kann keinen Dispatcher anbinden."
    }

    $sync = [hashtable]::Synchronized(@{
        LogBox     = $global:LogBox
        Dispatcher = $global:Window.Dispatcher
        LogFile    = $global:LogFile
        AppRoot    = $global:AppRoot
        Done       = $false
        Error      = $null
        Result     = $null
    })

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('Sync',       $sync)
    $rs.SessionStateProxy.SetVariable('UserScript', $ScriptBlock)
    foreach ($k in $Parameters.Keys) {
        $rs.SessionStateProxy.SetVariable($k, $Parameters[$k])
    }

    $ps = [powershell]::Create()
    $ps.Runspace = $rs

    [void]$ps.AddScript({

        function Write-Log {
            param(
                [string]$Message,
                [ValidateSet('Info','Success','Warning','Error','Debug')]
                [string]$Level = 'Info'
            )
            $stamp = (Get-Date).ToString('HH:mm:ss')
            $line  = "[$stamp] [$Level] $Message"
            try { Add-Content -Path $Sync.LogFile -Value $line -Encoding UTF8 } catch { }

            $hex = switch ($Level) {
                'Info'    { '#E0E0E0' }
                'Success' { '#4CAF50' }
                'Warning' { '#FFC107' }
                'Error'   { '#F44336' }
                'Debug'   { '#9E9E9E' }
                default   { '#E0E0E0' }
            }

            $action = {
                param($box, $text, $color)
                $brush = New-Object System.Windows.Media.SolidColorBrush(
                            [System.Windows.Media.ColorConverter]::ConvertFromString($color))
                $para  = New-Object System.Windows.Documents.Paragraph
                $para.Margin = '0'
                $run   = New-Object System.Windows.Documents.Run($text)
                $run.Foreground = $brush
                $para.Inlines.Add($run)
                $box.Document.Blocks.Add($para)
                $box.ScrollToEnd()
            }
            $Sync.Dispatcher.Invoke($action, @($Sync.LogBox, $line, $hex))
        }

        function Write-Host {
            param(
                [Parameter(Position=0, ValueFromPipeline=$true, ValueFromRemainingArguments=$true)]
                [object]$Object,
                [string]$ForegroundColor,
                [string]$BackgroundColor,
                [switch]$NoNewline,
                [string]$Separator = ' '
            )
            $text = if ($null -eq $Object) { '' } else { ($Object -join $Separator) }
            $level = switch ($ForegroundColor) {
                'Red'      { 'Error' }
                'DarkRed'  { 'Error' }
                'Yellow'   { 'Warning' }
                'Green'    { 'Success' }
                'Gray'     { 'Debug' }
                'DarkGray' { 'Debug' }
                default    { 'Info' }
            }
            Write-Log -Message $text -Level $level
        }

        function Read-Host {
            param([string]$Prompt = '')
            Write-Log "Read-Host unterdrueckt (Prompt: '$Prompt') - liefere leeren String." -Level Debug
            return ''
        }

        try {
            $Sync.Result = & $UserScript
        } catch {
            $Sync.Error = $_
            Write-Log ("Fehler im Runspace: " + $_.Exception.Message) -Level Error
        } finally {
            $Sync.Done = $true
        }
    })

    $handle = $ps.BeginInvoke()

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $timer.Add_Tick({
        if (-not $handle.IsCompleted) { return }
        $timer.Stop()
        try { [void]$ps.EndInvoke($handle) } catch { }
        $ps.Dispose()
        $rs.Close()
        $rs.Dispose()
        if ($OnComplete) {
            try { & $OnComplete $sync } catch {
                Write-Log ("OnComplete-Fehler: " + $_.Exception.Message) -Level Error
            }
        }
    }.GetNewClosure())
    $timer.Start()

    return $sync
}
