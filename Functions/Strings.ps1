# =============================================================================
#  Strings.ps1
#  Laedt die statischen UI-Texte aus Config\strings.de.json und schreibt sie
#  beim Start in die benannten XAML-Elemente (siehe Apply-Strings).
#
#  ==> PFLEGE / DOKU: doc\TEXTE_BEARBEITEN.md
#      Dort steht, wie man Texte aendert, neue hinzufuegt und welche Texte
#      bewusst NICHT ausgelagert sind (dynamische Laufzeit-Texte).
# =============================================================================

function Get-AppStringsPath {
    return (Join-Path $Script:AppRoot 'Config\strings.de.json')
}

function Import-AppStrings {
    [CmdletBinding()]
    param(
        [string]$Path = (Get-AppStringsPath)
    )

    if (-not (Test-Path $Path)) {
        throw "Strings-Datei nicht gefunden: $Path"
    }

    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    return ($raw | ConvertFrom-Json)
}

# === Mapping XAML-Element -> Text-Schluessel ===
# N = x:Name im XAML, P = Ziel-Property (Text/Content/Header),
# K = Schluessel in Config\strings.de.json.
# Mehrere Elemente duerfen denselben Schluessel verwenden (z. B. die drei
# "Ordner auswaehlen"-Buttons oder die "Bebilderte Hinweise"-Expander).
# ==> Neue Texte ergaenzen: erst hier eintragen, siehe doc\TEXTE_BEARBEITEN.md
$Script:StringMap = @(
    # Top-Bar
    @{ N = 'TxtBrandTitle';        P = 'Text';    K = 'Brand.Title' }
    @{ N = 'TxtBrandSubtitle';     P = 'Text';    K = 'Brand.Subtitle' }
    @{ N = 'LblTopKW';             P = 'Text';    K = 'Top.KW' }
    @{ N = 'LblTopToday';          P = 'Text';    K = 'Top.Today' }
    @{ N = 'LblTopSunday';         P = 'Text';    K = 'Top.LastSunday' }
    @{ N = 'LblTopWeekSelect';     P = 'Text';    K = 'Top.WeekSelect' }

    # Sidebar
    @{ N = 'LblSidebarSection';    P = 'Text';    K = 'Sidebar.Section' }
    @{ N = 'BtnStep1';             P = 'Content'; K = 'Sidebar.Step1' }
    @{ N = 'BtnStep2';             P = 'Content'; K = 'Sidebar.Step2' }
    @{ N = 'BtnStep3';             P = 'Content'; K = 'Sidebar.Step3' }
    @{ N = 'BtnStep4';             P = 'Content'; K = 'Sidebar.Step4' }
    @{ N = 'BtnStep5';             P = 'Content'; K = 'Sidebar.Step5' }
    @{ N = 'BtnSettings';          P = 'Content'; K = 'Sidebar.Settings' }

    # Schritt 1
    @{ N = 'Txt1Title';            P = 'Text';    K = 'Step1.Title' }
    @{ N = 'Txt1Desc';             P = 'Text';    K = 'Step1.Desc' }
    @{ N = 'Txt1AnleitungTitle';   P = 'Text';    K = 'Step1.AnleitungTitle' }
    @{ N = 'Txt1Anleitung1';       P = 'Text';    K = 'Step1.Anleitung.1' }
    @{ N = 'Txt1Anleitung2';       P = 'Text';    K = 'Step1.Anleitung.2' }
    @{ N = 'Txt1Anleitung3';       P = 'Text';    K = 'Step1.Anleitung.3' }
    @{ N = 'Btn1OpenPath';         P = 'Content'; K = 'Btn.Step1.OpenPath' }
    @{ N = 'Btn1Done';             P = 'Content'; K = 'Btn.Step1.Done' }

    # Schritt 2
    @{ N = 'Txt2Title';            P = 'Text';    K = 'Step2.Title' }
    @{ N = 'Txt2Desc';             P = 'Text';    K = 'Step2.Desc' }
    @{ N = 'Txt2OptionenTitle';    P = 'Text';    K = 'Step2.OptionenTitle' }
    @{ N = 'Chk2NurExcel';         P = 'Content'; K = 'Chk.Step2.NurExcel' }
    @{ N = 'Txt2PreviewTitle';     P = 'Text';    K = 'Step2.PreviewTitle' }
    @{ N = 'Txt2LastRunTitle';     P = 'Text';    K = 'Step2.LastRunTitle' }
    @{ N = 'Btn2Run';              P = 'Content'; K = 'Btn.Step2.Run' }
    @{ N = 'Btn2OpenPath';         P = 'Content'; K = 'Btn.Step2.OpenPath' }

    # Schritt 3
    @{ N = 'Txt3Title';            P = 'Text';    K = 'Step3.Title' }
    @{ N = 'Txt3Desc';             P = 'Text';    K = 'Step3.Desc' }
    @{ N = 'Txt3BezeichnungTitle'; P = 'Text';    K = 'Step3.BezeichnungTitle' }
    @{ N = 'Txt3FaelligkeitTitle'; P = 'Text';    K = 'Step3.FaelligkeitTitle' }
    @{ N = 'Btn3Copy';             P = 'Content'; K = 'Btn.Step3.Copy' }
    @{ N = 'Btn3Prosos';           P = 'Content'; K = 'Btn.Step3.Prosos' }
    @{ N = 'Txt3ChecklistTitle';   P = 'Text';    K = 'Step3.ChecklistTitle' }
    @{ N = 'Chk3Done0';            P = 'Content'; K = 'Chk.Step3.Done0' }
    @{ N = 'Chk3Done1';            P = 'Content'; K = 'Chk.Step3.Done1' }
    @{ N = 'Chk3Done2';            P = 'Content'; K = 'Chk.Step3.Done2' }
    @{ N = 'Chk3Done3';            P = 'Content'; K = 'Chk.Step3.Done3' }

    # Schritt 4
    @{ N = 'Txt4Title';            P = 'Text';    K = 'Step4.Title' }
    @{ N = 'Txt4Desc';             P = 'Text';    K = 'Step4.Desc' }
    @{ N = 'Txt4AufgabenTitle';    P = 'Text';    K = 'Step4.AufgabenTitle' }
    @{ N = 'Txt4AufgabenDesc';     P = 'Text';    K = 'Step4.AufgabenDesc' }
    @{ N = 'Txt4OrdnerTitle';      P = 'Text';    K = 'Step4.OrdnerTitle' }
    @{ N = 'Txt4StatusTitle';      P = 'Text';    K = 'Step4.StatusTitle' }
    @{ N = 'Btn4Run';              P = 'Content'; K = 'Btn.Step4.Run' }
    @{ N = 'Btn4OpenTask';         P = 'Content'; K = 'Btn.Step4.OpenTask' }

    # Schritt 5
    @{ N = 'Txt5Title';            P = 'Text';    K = 'Step5.Title' }
    @{ N = 'Txt5Desc';             P = 'Text';    K = 'Step5.Desc' }
    @{ N = 'Txt5LetzteTitle';      P = 'Text';    K = 'Step5.LetzteTitle' }
    @{ N = 'Chk5Done0';            P = 'Content'; K = 'Chk.Step5.Done0' }
    @{ N = 'Chk5Done1';            P = 'Content'; K = 'Chk.Step5.Done1' }
    @{ N = 'Chk5Done2';            P = 'Content'; K = 'Chk.Step5.Done2' }
    @{ N = 'Btn5MailTemplate';     P = 'Content'; K = 'Btn.Step5.MailTemplate' }
    @{ N = 'Btn5Finish';           P = 'Content'; K = 'Btn.Step5.Finish' }

    # Einstellungen
    @{ N = 'TxtSetTitle';          P = 'Text';    K = 'Settings.Title' }
    @{ N = 'TxtSetDesc';           P = 'Text';    K = 'Settings.Desc' }
    @{ N = 'TxtSetMsgTitle';       P = 'Text';    K = 'Settings.MsgTitle' }
    @{ N = 'TxtSetTaskTitle';      P = 'Text';    K = 'Settings.TaskTitle' }
    @{ N = 'TxtSetZahllaufTitle';  P = 'Text';    K = 'Settings.ZahllaufTitle' }
    @{ N = 'BtnSetMsgBrowse';      P = 'Content'; K = 'Btn.Settings.Browse' }
    @{ N = 'BtnSetTaskBrowse';     P = 'Content'; K = 'Btn.Settings.Browse' }
    @{ N = 'BtnSetZahllaufBrowse'; P = 'Content'; K = 'Btn.Settings.Browse' }
    @{ N = 'BtnSetSave';           P = 'Content'; K = 'Btn.Settings.Save' }
    @{ N = 'TxtSetBehaviorTitle';  P = 'Text';    K = 'Settings.BehaviorTitle' }
    @{ N = 'ChkSetSimulate';       P = 'Content'; K = 'Chk.Settings.Simulate' }
    @{ N = 'TxtSetSimulateDesc';   P = 'Text';    K = 'Settings.SimulateDesc' }

    # Rechtes Panel (Status)
    @{ N = 'LblRightSection';      P = 'Text';    K = 'Right.Section' }
    @{ N = 'LblRightProgress';     P = 'Text';    K = 'Right.Progress' }

    # Log
    @{ N = 'LblLogSection';        P = 'Text';    K = 'Log.Section' }
    @{ N = 'BtnOpenLog';           P = 'Content'; K = 'Btn.Log.Open' }
    @{ N = 'BtnClearLog';          P = 'Content'; K = 'Btn.Log.Clear' }

    # Bebilderte Hinweise (Expander-Header, fuenf mal derselbe Text)
    @{ N = 'Exp1Hints';            P = 'Header';  K = 'Hints.ExpanderHeader' }
    @{ N = 'Exp2Hints';            P = 'Header';  K = 'Hints.ExpanderHeader' }
    @{ N = 'Exp3Hints';            P = 'Header';  K = 'Hints.ExpanderHeader' }
    @{ N = 'Exp4Hints';            P = 'Header';  K = 'Hints.ExpanderHeader' }
    @{ N = 'Exp5Hints';            P = 'Header';  K = 'Hints.ExpanderHeader' }
)

# Setzt die statischen UI-Texte aus Config\strings.de.json in die benannten
# XAML-Elemente. Aufruf in Add_Loaded VOR Update-Context (Initialize-WeekPicker).
# Fehlt ein Schluessel oder ein Element, bleibt der XAML-Wert stehen und es
# wird eine Warnung ins Log geschrieben.
# ==> Doku: doc\TEXTE_BEARBEITEN.md
function Apply-Strings {
    $s = $Script:Strings
    if (-not $s) { Write-Log "Strings: keine Texte geladen." -Level Warning; return }

    # Window.Title separat (das Window selbst hat keinen x:Name).
    if ($s.PSObject.Properties['Window.Title']) {
        $Script:Window.Title = [string]$s.'Window.Title'
    }

    foreach ($m in $Script:StringMap) {
        if (-not $s.PSObject.Properties[$m.K]) {
            Write-Log "Strings: Schluessel fehlt -> '$($m.K)' (Element $($m.N))" -Level Warning
            continue
        }
        $el = $Script:Window.FindName($m.N)
        if (-not $el) {
            Write-Log "Strings: Element fehlt -> '$($m.N)' (Schluessel $($m.K))" -Level Warning
            continue
        }
        $el.($m.P) = [string]$s.($m.K)
    }
    Write-Log "UI-Texte aus strings.de.json angewendet." -Level Debug
}
