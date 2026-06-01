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
. (Join-Path $Script:AppRoot 'Functions\Strings.ps1')

# -- Module laden --------------------------------------------------------------
Get-ChildItem -Path (Join-Path $Script:AppRoot 'Modules') -Filter '*.psm1' |
    ForEach-Object { Import-Module $_.FullName -Force -DisableNameChecking }

# -- Config laden --------------------------------------------------------------
$Script:Config = Import-AppConfig
$global:Config = $Script:Config

# -- Statische UI-Texte laden (siehe doc\TEXTE_BEARBEITEN.md) -------------------
$Script:Strings = Import-AppStrings
$global:Strings = $Script:Strings

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
    CmbWeek          = Find-Element 'CmbWeek'

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
    Btn1OpenOutlook  = Find-Element 'Btn1OpenOutlook'
    Btn1OpenPath     = Find-Element 'Btn1OpenPath'
    Btn1CopyPath     = Find-Element 'Btn1CopyPath'
    Btn1Done         = Find-Element 'Btn1Done'
    Sp1Hints         = Find-Element 'Sp1Hints'

    # Step 2
    Chk2NurExcel     = Find-Element 'Chk2NurExcel'
    Txt2Preview      = Find-Element 'Txt2Preview'
    Txt2LastRun      = Find-Element 'Txt2LastRun'
    Btn2Run          = Find-Element 'Btn2Run'
    Btn2OpenExcel    = Find-Element 'Btn2OpenExcel'
    Btn2OpenPath     = Find-Element 'Btn2OpenPath'
    Txt2MergeStatus  = Find-Element 'Txt2MergeStatus'
    Btn2MergePdf     = Find-Element 'Btn2MergePdf'
    Btn2OpenMergedPdf = Find-Element 'Btn2OpenMergedPdf'
    Btn2Done         = Find-Element 'Btn2Done'
    Sp2Hints         = Find-Element 'Sp2Hints'

    # Step 3
    Txt3Bezeichnung  = Find-Element 'Txt3Bezeichnung'
    Dp3Faelligkeit   = Find-Element 'Dp3Faelligkeit'
    Btn3Copy         = Find-Element 'Btn3Copy'
    Btn3Prosos       = Find-Element 'Btn3Prosos'
    Btn3Summe        = Find-Element 'Btn3Summe'
    Txt3Summe        = Find-Element 'Txt3Summe'
    Btn3Done         = Find-Element 'Btn3Done'
    Sp3Hints         = Find-Element 'Sp3Hints'

    # Step 4
    Txt4Ordner       = Find-Element 'Txt4Ordner'
    Txt4Status       = Find-Element 'Txt4Status'
    Btn4Run          = Find-Element 'Btn4Run'
    Btn4OpenTask     = Find-Element 'Btn4OpenTask'
    Btn4OpenXlsx     = Find-Element 'Btn4OpenXlsx'
    Btn4OpenPdf      = Find-Element 'Btn4OpenPdf'
    Btn4Done         = Find-Element 'Btn4Done'
    Sp4Tasks         = Find-Element 'Sp4Tasks'
    Txt4TaskStatus   = Find-Element 'Txt4TaskStatus'
    Sp4Hints         = Find-Element 'Sp4Hints'

    # Step 5
    Btn5MailTemplate = Find-Element 'Btn5MailTemplate'
    Btn5DeleteTemp   = Find-Element 'Btn5DeleteTemp'
    Btn5Finish       = Find-Element 'Btn5Finish'
    Sp5Hints         = Find-Element 'Sp5Hints'

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
function Initialize-WeekPicker {
    # Liste der letzten Wochen aufbauen, Default = vergangene Woche (Index 0).
    $Script:Weeks = Get-RecentWeeks -From (Get-Date) -Count 4
    $ui.CmbWeek.Items.Clear()
    foreach ($w in $Script:Weeks) { [void]$ui.CmbWeek.Items.Add($w.Display) }
    $ui.CmbWeek.SelectedIndex = 0   # loest SelectionChanged -> Update-Context aus
}

function Update-Context {
    if (-not $Script:Weeks) { return }
    $idx = $ui.CmbWeek.SelectedIndex
    if ($idx -lt 0) { return }       # transienter Zustand beim Neuaufbau

    $refDate = $Script:Weeks[$idx].Sunday
    $ctx = Get-ZahllaufContext -RefDate $refDate
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

# Erzeugt pro Config.ScheduledTasks-Eintrag einen Button in Sp4Tasks
function Initialize-ScheduledTaskButtons {
    if (-not $ui.Sp4Tasks) { return }
    $tasks = $Script:Config.ScheduledTasks
    if (-not $tasks) { return }

    # In lokale Variablen heben, damit GetNewClosure() sie per Wert einfaengt.
    # $Script:-Referenzen wuerden im Closure-Modulscope auf $null aufloesen.
    $btnStyle     = $Script:Window.Resources['BtnPrimary']
    $statusBlock  = $ui.Txt4TaskStatus
    $brushSuccess = $Script:Window.Resources['Success']
    $brushDanger  = $Script:Window.Resources['Danger']

    $ui.Sp4Tasks.Children.Clear()
    foreach ($t in $tasks) {
        $btn         = New-Object System.Windows.Controls.Button
        $btn.Content = [string]$t.Label
        $btn.Style   = $btnStyle
        $btn.Margin  = [System.Windows.Thickness]::new(0, 0, 10, 0)
        $taskId      = [string]$t.Id
        $handler = {
            Invoke-ScheduledTaskById `
                -Id              $taskId `
                -StatusTextBlock $statusBlock `
                -BrushSuccess    $brushSuccess `
                -BrushDanger     $brushDanger
        }.GetNewClosure()
        $btn.Add_Click($handler)
        [void]$ui.Sp4Tasks.Children.Add($btn)
    }
    Write-Log "Aufgabenplanung: $(@($tasks).Count) Task(s) geladen." -Level Debug
}

# =============================================================================
#  Bebilderte Hinweise (Galerie)
# =============================================================================
# Zeigt ein Hinweisbild gross in einem modalen Fenster (oder Platzhalter-Hinweis).
function global:Show-HintImage {
    param([string]$Path, [string]$Caption, $Owner)

    if (-not [System.IO.File]::Exists($Path)) {
        [void][System.Windows.MessageBox]::Show(
            "Platzhalter - fuer diesen Hinweis ist noch kein Bild hinterlegt.`r`n`r`n$Caption",
            "Bebilderter Hinweis")
        return
    }

    $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
    $bmp.BeginInit()
    $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bmp.UriSource   = New-Object System.Uri($Path)
    $bmp.EndInit()

    $img = New-Object System.Windows.Controls.Image
    $img.Stretch = [System.Windows.Media.Stretch]::Uniform
    $img.Source  = $bmp
    $img.Margin  = [System.Windows.Thickness]::new(10)

    $win = New-Object System.Windows.Window
    $win.Title  = $Caption
    $win.Width  = 960
    $win.Height = 720
    $win.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    $win.Background = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(30, 30, 46))
    if ($Owner) { $win.Owner = $Owner }
    $win.Content = $img
    [void]$win.ShowDialog()
}

# Baut pro Step eine Thumbnail-Galerie aus Config.Hints. Fehlt eine Bilddatei,
# erscheint eine Platzhalter-Kachel. Klick oeffnet die Grossansicht.
function Initialize-HintGalleries {
    $hints = $Script:Config.Hints
    if (-not $hints) { return }

    $panels = @{
        Step1 = $ui.Sp1Hints
        Step2 = $ui.Sp2Hints
        Step3 = $ui.Sp3Hints
        Step4 = $ui.Sp4Hints
        Step5 = $ui.Sp5Hints
    }

    # In lokale Variablen heben (Closure-Scope, siehe CLAUDE.MD).
    $brushBorder   = $Script:Window.Resources['Border']
    $brushTertiary = $Script:Window.Resources['BgTertiary']
    $brushMuted    = $Script:Window.Resources['TextMuted']
    $brushText     = $Script:Window.Resources['TextPrimary']
    $owner         = $Script:Window

    foreach ($key in $panels.Keys) {
        $panel = $panels[$key]
        if (-not $panel) { continue }
        $panel.Children.Clear()

        $items = $hints.$key
        if (-not $items) { continue }

        foreach ($h in $items) {
            $imgPath = Join-Path $global:AppRoot ([string]$h.Image)
            $caption = [string]$h.Caption
            $exists  = [System.IO.File]::Exists($imgPath)

            $inner = New-Object System.Windows.Controls.StackPanel

            if ($exists) {
                $bmp = New-Object System.Windows.Media.Imaging.BitmapImage
                $bmp.BeginInit()
                $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
                $bmp.UriSource   = New-Object System.Uri($imgPath)
                $bmp.EndInit()
                $thumb = New-Object System.Windows.Controls.Image
                $thumb.Height  = 110
                $thumb.Stretch = [System.Windows.Media.Stretch]::Uniform
                $thumb.Source  = $bmp
                [void]$inner.Children.Add($thumb)
            } else {
                $ph = New-Object System.Windows.Controls.TextBlock
                $ph.Text                = "Bild folgt"
                $ph.Foreground          = $brushMuted
                $ph.Height              = 110
                $ph.TextAlignment       = [System.Windows.TextAlignment]::Center
                $ph.Padding             = [System.Windows.Thickness]::new(0, 44, 0, 0)
                [void]$inner.Children.Add($ph)
            }

            $cap = New-Object System.Windows.Controls.TextBlock
            $cap.Text          = $caption
            $cap.Foreground    = $brushText
            $cap.FontSize      = 11
            $cap.TextWrapping  = [System.Windows.TextWrapping]::Wrap
            $cap.TextAlignment = [System.Windows.TextAlignment]::Center
            $cap.Margin        = [System.Windows.Thickness]::new(6, 8, 6, 2)
            [void]$inner.Children.Add($cap)

            $tile = New-Object System.Windows.Controls.Border
            $tile.Width           = 190
            $tile.Margin          = [System.Windows.Thickness]::new(0, 0, 12, 12)
            $tile.Padding         = [System.Windows.Thickness]::new(6)
            $tile.CornerRadius    = [System.Windows.CornerRadius]::new(6)
            $tile.BorderThickness = [System.Windows.Thickness]::new(1)
            $tile.BorderBrush     = $brushBorder
            $tile.Background      = $brushTertiary
            $tile.Cursor          = [System.Windows.Input.Cursors]::Hand
            $tile.Child           = $inner

            $pathLocal  = $imgPath
            $capLocal   = $caption
            $ownerLocal = $owner
            $handler = {
                Show-HintImage -Path $pathLocal -Caption $capLocal -Owner $ownerLocal
            }.GetNewClosure()
            $tile.Add_MouseLeftButtonUp($handler)

            [void]$panel.Children.Add($tile)
        }
    }
    Write-Log "Bebilderte Hinweise geladen." -Level Debug
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
$ui.CmbWeek.Add_SelectionChanged({ Update-Context })
$ui.BtnClearLog.Add_Click({ Clear-Log; Write-Log "Log geleert." -Level Debug })
$ui.BtnOpenLog.Add_Click({ [System.Diagnostics.Process]::Start('notepad.exe', $Script:LogFile) | Out-Null })

# Step 1
$ui.Btn1OpenOutlook.Add_Click({ Invoke-Step1OpenOutlook })
$ui.Btn1OpenPath.Add_Click({ Invoke-Step1OpenPath -Path $Script:Config.Paths.TempRoot })
$ui.Btn1CopyPath.Add_Click({ Invoke-Step1CopyPath -Path $Script:Config.Paths.TempRoot })
$ui.Btn1Done.Add_Click({ Invoke-Step1MarkDone })

# Step 2
$ui.Btn2Run.Add_Click({
    $ui.Txt2LastRun.Text = "Letzter Aufruf: $((Get-Date).ToString('dd.MM.yyyy HH:mm')) - laeuft..."
    Invoke-Step2Run `
        -NurExcel:$ui.Chk2NurExcel.IsChecked `
        -StatusTextBlock $ui.Txt2LastRun `
        -BrushSuccess $Script:Window.Resources['Success'] `
        -BrushDanger  $Script:Window.Resources['Danger']
})
$ui.Btn2OpenExcel.Add_Click({ Invoke-Step2OpenExcel -Path (Join-Path $Script:Config.Paths.TaskFolder 'alleRechnungen.xlsx') })
$ui.Btn2OpenPath.Add_Click({ Invoke-Step2OpenPath -Path $Script:Config.Paths.ExtractedFolder })
$ui.Btn2MergePdf.Add_Click({
    $ui.Txt2MergeStatus.Text = "Laeuft..."
    Invoke-Step2MergePdf `
        -StatusTextBlock $ui.Txt2MergeStatus `
        -BrushSuccess $Script:Window.Resources['Success'] `
        -BrushDanger  $Script:Window.Resources['Danger']
})
$ui.Btn2OpenMergedPdf.Add_Click({ Invoke-Step2OpenMergedPdf -Path (Join-Path $Script:Config.Paths.TaskFolder 'alleRechnungen_Anhaenge.pdf') })
$ui.Btn2Done.Add_Click({ Invoke-Step2MarkDone })

# Step 3
$ui.Btn3Copy.Add_Click({ Invoke-Step3Copy -Text $ui.Txt3Bezeichnung.Text })
$ui.Btn3Prosos.Add_Click({ Invoke-Step3OpenProsos })
$ui.Btn3Done.Add_Click({ Invoke-Step3MarkDone })
$ui.Btn3Summe.Add_Click({
    Invoke-Step3ReadSumme `
        -TaskFolder     $Script:Config.Paths.TaskFolder `
        -SummeTextBlock $ui.Txt3Summe `
        -BrushSuccess   $Script:Window.Resources['Success'] `
        -BrushDanger    $Script:Window.Resources['Danger']
})

# Step 4
$ui.Btn4Run.Add_Click({
    Invoke-Step4Run `
        -StatusTextBlock $ui.Txt4Status `
        -BrushSuccess $Script:Window.Resources['Success'] `
        -BrushDanger  $Script:Window.Resources['Danger'] `
        -BtnOpenXlsx $ui.Btn4OpenXlsx `
        -BtnOpenPdf  $ui.Btn4OpenPdf
})
$ui.Btn4OpenTask.Add_Click({ Invoke-Step4OpenTask -Path $Script:Config.Paths.TaskFolder })
$ui.Btn4OpenXlsx.Add_Click({ Invoke-Step4OpenAbgleichFile -Folder ([string]$ui.Btn4OpenXlsx.Tag) -Pattern '*alleRechnungen.xlsx' })
$ui.Btn4OpenPdf.Add_Click({  Invoke-Step4OpenAbgleichFile -Folder ([string]$ui.Btn4OpenPdf.Tag)  -Pattern '*ZahllisteHHSTGesamtBetr.pdf' })
$ui.Btn4Done.Add_Click({ Invoke-Step4MarkDone })

# Step 5
$ui.Btn5MailTemplate.Add_Click({
    $summe = $ui.Txt3Summe.Text
    if ([string]::IsNullOrWhiteSpace($summe) -or $summe -eq '...') { $summe = '(nicht ermittelt)' }
    $zielpfad = Join-Path $Script:Config.Paths.ZahllaufFolder $Script:Context.OrdnerName
    Invoke-Step5MailTemplate `
        -Bezeichnung  $Script:Context.Bezeichnung `
        -Summe        $summe `
        -ZielpfadLink $zielpfad
})
$ui.Btn5DeleteTemp.Add_Click({ Invoke-Step5DeleteTemp })
$ui.Btn5Finish.Add_Click({ Invoke-Step5Finish -Bezeichnung $Script:Context.Bezeichnung })

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
    Apply-Strings   # statische UI-Texte VOR Update-Context setzen
    Initialize-WeekPicker
    Initialize-ScheduledTaskButtons
    Initialize-HintGalleries
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
