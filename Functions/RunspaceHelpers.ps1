# =============================================================================
#  RunspaceHelpers.ps1
#  Fuehrt langlaufende Skripte in einem eigenen STA-Runspace aus, damit die
#  WPF-GUI responsiv bleibt. Write-Host / Write-Log aus dem Runspace werden
#  per Dispatcher in die globale Log-Box (Script:LogBox) umgeleitet.
# =============================================================================

function Start-RunspaceJob {
    [CmdletBinding()]
    param(
        # Eigentlicher Code, der im Runspace laeuft.
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        # Variablen, die in den Runspace gesetzt werden (Name -> Wert).
        [hashtable]$Parameters = @{},

        # Wird im GUI-Thread aufgerufen, sobald der Runspace fertig ist.
        # Bekommt das Sync-Hashtable (mit .Error, .Result, .Done).
        [scriptblock]$OnComplete
    )

    if (-not $Script:Window) {
        throw "Start-RunspaceJob: Script:Window fehlt - kann keinen Dispatcher anbinden."
    }

    # --- gemeinsamer Zustand zwischen GUI-Thread und Runspace --------------
    $sync = [hashtable]::Synchronized(@{
        LogBox     = $Script:LogBox
        Dispatcher = $Script:Window.Dispatcher
        LogFile    = $Script:LogFile
        AppRoot    = $Script:AppRoot
        Done       = $false
        Error      = $null
        Result     = $null
    })

    # --- Runspace + PowerShell-Pipeline ------------------------------------
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

    # Der Wrapper definiert Log-Funktionen im Runspace und ruft dann das
    # eigentliche User-Skript auf.
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

        # Write-Host der Originalskripte umlenken (anhand der Farbe das
        # passende Log-Level waehlen).
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

        # Read-Host nicht-interaktiv machen: Defaultantwort ist leer
        # -> die Originalskripte nehmen den jeweiligen Standardpfad.
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

    # --- Abschluss-Erkennung via DispatcherTimer ---------------------------
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
