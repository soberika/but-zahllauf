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
    $rs.SessionStateProxy.SetVariable('Sync',           $sync)
    $rs.SessionStateProxy.SetVariable('UserScriptText', $ScriptBlock.ToString())
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
            $stamp = [DateTime]::Now.ToString('HH:mm:ss')
            $line  = "[$stamp] [$Level] $Message"
            try {
                [System.IO.File]::AppendAllText(
                    $Sync.LogFile,
                    ($line + "`r`n"),
                    [System.Text.Encoding]::UTF8)
            } catch { }

            $hex = switch ($Level) {
                'Info'    { '#E0E0E0' }
                'Success' { '#4CAF50' }
                'Warning' { '#FFC107' }
                'Error'   { '#F44336' }
                'Debug'   { '#9E9E9E' }
                default   { '#E0E0E0' }
            }

            $box = $Sync.LogBox
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

            [void]$Sync.Dispatcher.Invoke([Action]$append)
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
            # Script-Block im Runspace neu erzeugen, damit er die inneren
            # Write-Log/Write-Host/Read-Host-Funktionen aufloest.
            $userSb = [scriptblock]::Create($UserScriptText)
            $Sync.Result = & $userSb
        } catch {
            $Sync.Error = $_
            try { Write-Log ("Fehler im Runspace: " + $_.Exception.Message) -Level Error } catch { }
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

        # WICHTIG: OnComplete VOR Dispose ausfuehren.
        if ($OnComplete) {
            try { & $OnComplete $sync } catch {
                try { Write-Log ("OnComplete-Fehler: " + $_.Exception.Message) -Level Error } catch { }
            }
        }

        try { $ps.Dispose() } catch { }
        try { $rs.Close() } catch { }
        try { $rs.Dispose() } catch { }
    }.GetNewClosure())
    $timer.Start()

    return $sync
}
