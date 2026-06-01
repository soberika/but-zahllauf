<#
.SYNOPSIS
    viafintech Zahllauf Dashboard - Hauptskript (Phase 1, Grundgeruest).

.DESCRIPTION
    Modulare WPF-GUI fuer den viafintech-Zahllaufprozess.
    Phase 1: nur Grundgeruest (UI, Navigation, Auto-Berechnungen, Logging,
    Platzhalter-Buttons). Echte Integration der bestehenden .ps1-Skripte
    erfolgt in Phase 2.

.NOTES
    PowerShell 5.1 (Desktop), STA-Apartment fuer WPF.
    Start: .\Main.ps1
#>

#Requires -Version 5.1

# -- STA pruefen ---------------------------------------------------------------
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    Write-Host "Starte neu im STA-Modus..." -ForegroundColor Yellow
    powershell.exe -NoProfile -Sta -File $PSCommandPath @args
    return
}

# -- Pfade ---------------------------------------------------------------------
$Script:AppRoot = Split-Path -Parent $PSCommandPath
$global:AppRoot = $Script:AppRoot
$Script:LogDir  = Join-Path $Script:AppRoot 'Logs'
if (-not (Test-Path $Script:LogDir)) { New-Item -ItemType Directory -Path $Script:LogDir | Out-Null }
$Script:LogFile = Join-Path $Script:LogDir ("zahllauf_{0:yyyy-MM-dd}.log" -f (Get-Date))
$global:LogFile = $Script:LogFile

# -- Assemblies ----------------------------------------------------------------
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms   # nur fuer Set-Clipboard-Fallback

# -- Functions laden -----------------------------------------------------------
. (Join-Path $Script:AppRoot 'Functions\DateHelpers.ps1')
. (Join-Path $Script:AppRoot 'Functions\Logging.ps1')
. (Join-Path $Script:AppRoot 'Functions\ConfigLoader.ps1')
. (Join-Path $Script:AppRoot 'Functions\RunspaceHelpers.ps1')

# -- Module laden --------------------------------------------------------------
Get-ChildItem -Path (Join-Path $Script:AppRoot 'Modules') -Filter '*.psm1' |
    ForEach-Object { Import-Module $_.FullName -Force -DisableNameChecking }

# -- Config laden --------------------------------------------------------------
$Script:Config = Import-AppConfig
$global:Config = $Script:Config

# =============================================================================
#  XAML laden
# =============================================================================
[xml]$xaml = Get-Content -Path (Join-Path $Script:AppRoot 'Views\MainWindow.xaml') -Raw -Encoding UTF8
$reader    = New-Object System.Xml.XmlNodeReader $xaml
$Script:Window = [Windows.Markup.XamlReader]::Load($reader)
$global:Window = $Script:Window

# -- Hilfsfunktion: Element nach Name suchen -----------------------------------
function Find-Element {
    param([string]$Name)
    return $Script:Window.FindName($Name)
}

# =============================================================================
#  Element-Referenzen sammeln
# =============================================================================
$ui = @{
    # Top-Bar
    TxtKW            = Find-Element 'TxtKW'
    TxtDate          = Find-Element 'TxtDate'
    TxtSunday        = Find-Element 'TxtSunday'

    # Right panel
    TxtKWRight       = Find-Element 'TxtKWRight'
    TxtYearRight     = Find-Element 'TxtYearRight'
    TxtBezRight      = Find-Element 'TxtBezRight'
    TxtOrdRight      = Find-Element 'TxtOrdRight'

    # Sidebar
    BtnStep1         = Find-Element 'BtnStep1'
    BtnStep2         = Find-Element 'BtnStep2'
    BtnStep3         = Find-Element 'BtnStep3'
    BtnStep4         = Find-Element 'BtnStep4'
    BtnStep5         = Find-Element 'BtnStep5'

    # Pages
    PageStep1        = Find-Element 'PageStep1'
    PageStep2        = Find-Element 'PageStep2'
    PageStep3        = Find-Element 'PageStep3'
    PageStep4        = Find-Element 'PageStep4'
    PageStep5        = Find-Element 'PageStep5'

    # Step 1
    Btn1OpenPath     = Find-Element 'Btn1OpenPath'
    Btn1Done         = Find-Element 'Btn1Done'

    # Step 2
    Chk2NurExcel     = Find-Element 'Chk2NurExcel'
    Chk2NurExcelCsv  = Find-Element 'Chk2NurExcelCsv'
    Txt2Preview      = Find-Element 'Txt2Preview'
    Txt2LastRun      = Find-Element 'Txt2LastRun'
    Btn2Run          = Find-Element 'Btn2Run'
    Btn2OpenPath     = Find-Element 'Btn2OpenPath'

    # Step 3
    Txt3Bezeichnung  = Find-Element 'Txt3Bezeichnung'
    Dp3Faelligkeit   = Find-Element 'Dp3Faelligkeit'
    Btn3Copy         = Find-Element 'Btn3Copy'
    Btn3Prosos       = Find-Element 'Btn3Prosos'

    # Step 4
    Txt4Ordner       = Find-Element 'Txt4Ordner'
    Txt4Status       = Find-Element 'Txt4Status'
    Btn4Run          = Find-Element 'Btn4Run'
    Btn4Check        = Find-Element 'Btn4Check'
    Btn4OpenTask     = Find-Element 'Btn4OpenTask'

    # Step 5
    Btn5MailTemplate = Find-Element 'Btn5MailTemplate'
    Btn5Finish       = Find-Element 'Btn5Finish'

    # Log
    LogBox           = Find-Element 'LogBox'
    BtnClearLog      = Find-Element 'BtnClearLog'
    BtnOpenLog       = Find-Element 'BtnOpenLog'

    # Settings
    BtnSettings          = Find-Element 'BtnSettings'
    PageSettings         = Find-Element 'PageSettings'
    TxtSetMsg            = Find-Element 'TxtSetMsg'
    TxtSetTask           = Find-Element 'TxtSetTask'
    TxtSetZahllauf       = Find-Element 'TxtSetZahllauf'
    BtnSetMsgBrowse      = Find-Element 'BtnSetMsgBrowse'
    BtnSetTaskBrowse     = Find-Element 'BtnSetTaskBrowse'
    BtnSetZahllaufBrowse = Find-Element 'BtnSetZahllaufBrowse'
    BtnSetSave           = Find-Element 'BtnSetSave'
    TxtSetStatus         = Find-Element 'TxtSetStatus'
    ChkSetSimulate       = Find-Element 'ChkSetSimulate'
}

$Script:LogBox = $ui.LogBox
$global:LogBox = $Script:LogBox

# =============================================================================
#  Navigation
# =============================================================================
$Script:Pages = @{
    1 = $ui.PageStep1
    2 = $ui.PageStep2
    3 = $ui.PageStep3
    4 = $ui.PageStep4
    5 = $ui.PageStep5
    6 = $ui.PageSettings
}
$Script:SideButtons = @{
    1 = $ui.BtnStep1
    2 = $ui.BtnStep2
    3 = $ui.BtnStep3
    4 = $ui.BtnStep4
    5 = $ui.BtnStep5
    6 = $ui.BtnSettings
}
$Script:CurrentStep = 1

function Show-Step {
    param([int]$Id)

    foreach ($k in $Script:Pages.Keys) {
        $Script:Pages[$k].Visibility = if ($k -eq $Id) { 'Visible' } else { 'Collapsed' }
    }

    $styleActive = $Script:Window.Resources['SideBtnActive']
    $styleNormal = $Script:Window.Resources['SideBtn']
    foreach ($k in $Script:SideButtons.Keys) {
        $Script:SideButtons[$k].Style = if ($k -eq $Id) { $styleActive } else { $styleNormal }
    }

    $Script:CurrentStep = $Id
    Write-Log "Navigation -> Schritt $Id" -Level Debug
}

# =============================================================================
#  Werte berechnen + UI aktualisieren
# =============================================================================
function Update-Context {
    $ctx = Get-ZahllaufContext
    $Script:Context = $ctx

    # Top-Bar
    $ui.TxtKW.Text     = "KW $($ctx.KW)"
    $ui.TxtDate.Text   = "$($ctx.Date) ($($ctx.Weekday))"
    $ui.TxtSunday.Text = $ctx.LastSundayStr

    # Right panel
    $ui.TxtKWRight.Text   = "KW $($ctx.KW)"
    $ui.TxtYearRight.Text = "$($ctx.Year)"
    $ui.TxtBezRight.Text  = $ctx.Bezeichnung
    $ui.TxtOrdRight.Text  = $ctx.OrdnerName

    # Step 2
    $ui.Txt2Preview.Text  = $ctx.Bezeichnung

    # Step 3
    $ui.Txt3Bezeichnung.Text   = $ctx.Bezeichnung
    $ui.Dp3Faelligkeit.SelectedDate = $ctx.LastSunday

    # Step 4
    $ui.Txt4Ordner.Text = $ctx.OrdnerName

    Write-Log "Kontext aktualisiert: KW=$($ctx.KW), Bezeichnung='$($ctx.Bezeichnung)'" -Level Info
}

# =============================================================================
#  Event-Handler
# =============================================================================
# Sidebar
$ui.BtnStep1.Add_Click({ Show-Step 1 })
$ui.BtnStep2.Add_Click({ Show-Step 2 })
$ui.BtnStep3.Add_Click({ Show-Step 3 })
$ui.BtnStep4.Add_Click({ Show-Step 4 })
$ui.BtnStep5.Add_Click({ Show-Step 5 })

# Top-Bar
$ui.BtnClearLog.Add_Click({ Clear-Log; Write-Log "Log geleert." -Level Debug })
$ui.BtnOpenLog.Add_Click({ [System.Diagnostics.Process]::Start('notepad.exe', $Script:LogFile) | Out-Null })

# Step 1
$ui.Btn1OpenPath.Add_Click({ Invoke-Step1OpenPath -Path $Script:Config.Paths.TempRoot })
$ui.Btn1Done.Add_Click({ Invoke-Step1MarkDone })

# Step 2
$ui.Btn2Run.Add_Click({
    $ui.Txt2LastRun.Text = "Letzter Aufruf: $((Get-Date).ToString('dd.MM.yyyy HH:mm')) - laeuft..."
    Invoke-Step2Run `
        -NurExcel:$ui.Chk2NurExcel.IsChecked `
        -NurExcelCsv:$ui.Chk2NurExcelCsv.IsChecked `
        -StatusTextBlock $ui.Txt2LastRun `
        -BrushSuccess $Script:Window.Resources['Success'] `
        -BrushDanger  $Script:Window.Resources['Danger']
})
$ui.Btn2OpenPath.Add_Click({ Invoke-Step2OpenPath -Path $Script:Config.Paths.TempRoot })

# Step 3
$ui.Btn3Copy.Add_Click({ Invoke-Step3Copy -Text $ui.Txt3Bezeichnung.Text })
$ui.Btn3Prosos.Add_Click({ Invoke-Step3OpenProsos })

# Step 4
$ui.Btn4Run.Add_Click({
    Invoke-Step4Run `
        -StatusTextBlock $ui.Txt4Status `
        -BrushSuccess $Script:Window.Resources['Success'] `
        -BrushDanger  $Script:Window.Resources['Danger']
})
$ui.Btn4Check.Add_Click({ Invoke-Step4CheckCorrection })
$ui.Btn4OpenTask.Add_Click({ Invoke-Step4OpenTask -Path $Script:Config.Paths.TaskFolder })

# Step 5
$ui.Btn5MailTemplate.Add_Click({ Invoke-Step5MailTemplate -Bezeichnung $Script:Context.Bezeichnung })
$ui.Btn5Finish.Add_Click({ Invoke-Step5Finish })

# Settings
$ui.BtnSettings.Add_Click({ Show-Step 6 })

function Select-FolderInto {
    param($TextBox, [string]$Description)
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = $Description
    if (-not [string]::IsNullOrWhiteSpace($TextBox.Text) -and (Test-Path $TextBox.Text)) {
        $dlg.SelectedPath = $TextBox.Text
    }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TextBox.Text = $dlg.SelectedPath
    }
}

$ui.BtnSetMsgBrowse.Add_Click({      Select-FolderInto -TextBox $ui.TxtSetMsg      -Description 'Msg-Ordner waehlen' })
$ui.BtnSetTaskBrowse.Add_Click({     Select-FolderInto -TextBox $ui.TxtSetTask     -Description 'Task-Ordner waehlen' })
$ui.BtnSetZahllaufBrowse.Add_Click({ Select-FolderInto -TextBox $ui.TxtSetZahllauf -Description 'Zahllauf-Ordner waehlen' })

$ui.BtnSetSave.Add_Click({
    $Script:Config.Paths.MsgFolder           = $ui.TxtSetMsg.Text
    $Script:Config.Paths.TaskFolder          = $ui.TxtSetTask.Text
    $Script:Config.Paths.ZahllaufFolder      = $ui.TxtSetZahllauf.Text
    $Script:Config.Behavior.SimulateCleanup  = [bool]$ui.ChkSetSimulate.IsChecked
    Save-AppConfig
    $ui.TxtSetStatus.Text = "Einstellungen gespeichert ($((Get-Date).ToString('HH:mm:ss')))."
    $ui.TxtSetStatus.Foreground = $Script:Window.Resources['Success']
    Write-Log "Pfade gespeichert: Msg='$($Script:Config.Paths.MsgFolder)' Task='$($Script:Config.Paths.TaskFolder)' Zahllauf='$($Script:Config.Paths.ZahllaufFolder)'" -Level Info
})

# =============================================================================
#  Ungefangene Fehler -> Logdatei
# =============================================================================
[AppDomain]::CurrentDomain.add_UnhandledException({
    param($sender, $e)
    $ex  = $e.ExceptionObject
    $msg = "$([DateTime]::Now.ToString('HH:mm:ss')) [FATAL] $($ex.GetType().Name): $($ex.Message)`r`n$($ex.StackTrace)`r`n"
    try { [System.IO.File]::AppendAllText($global:LogFile, $msg, [System.Text.Encoding]::UTF8) } catch { }
})

# =============================================================================
#  Initialer Lauf
# =============================================================================
$Script:Window.Add_Loaded({
    Write-Log "viafintech Zahllauf Dashboard gestartet." -Level Success
    Write-Log ("PS {0}  |  OS {1}  |  Log: {2}" -f `
        $PSVersionTable.PSVersion, `
        [System.Environment]::OSVersion.VersionString, `
        $Script:LogFile) -Level Debug
    Update-Context
    $ui.TxtSetMsg.Text           = [string]$Script:Config.Paths.MsgFolder
    $ui.TxtSetTask.Text          = [string]$Script:Config.Paths.TaskFolder
    $ui.TxtSetZahllauf.Text      = [string]$Script:Config.Paths.ZahllaufFolder
    $ui.ChkSetSimulate.IsChecked = [bool]$Script:Config.Behavior.SimulateCleanup
    Show-Step 1
})

# =============================================================================
#  Show Window
# =============================================================================
[void]$Script:Window.ShowDialog()

# Ungefangene PS-Fehler aus der Session nachtraeglich sichern
if ($Error.Count -gt 0) {
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("$([DateTime]::Now.ToString('HH:mm:ss')) [SESSION-ERRORS] $($Error.Count) ungefangene Fehler:")
    foreach ($e in $Error) {
        [void]$sb.AppendLine("  [$($e.InvocationInfo.ScriptName):$($e.InvocationInfo.ScriptLineNumber)] $($e.Exception.Message)")
    }
    try { [System.IO.File]::AppendAllText($Script:LogFile, $sb.ToString(), [System.Text.Encoding]::UTF8) } catch { }
}
