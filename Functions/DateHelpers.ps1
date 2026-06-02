# =============================================================================
#  DateHelpers.ps1
#  Datumsfunktionen fuer den viafintech Zahllauf
# =============================================================================

function global:Get-IsoCalendarWeek {
    <#
    .SYNOPSIS
        Liefert die ISO-8601-Kalenderwoche (Montag = Wochenbeginn).
    #>
    param([datetime]$Date = (Get-Date))

    $cal  = [System.Globalization.CultureInfo]::InvariantCulture.Calendar
    $rule = [System.Globalization.CalendarWeekRule]::FirstFourDayWeek
    $dow  = [System.DayOfWeek]::Monday
    return $cal.GetWeekOfYear($Date, $rule, $dow)
}

function Get-IsoYear {
    <#
    .SYNOPSIS
        Liefert das ISO-Jahr (kann an Jahreswechseln vom Kalenderjahr abweichen).
    #>
    param([datetime]$Date = (Get-Date))

    # ISO-Jahr: Donnerstag der gleichen Woche bestimmt das Jahr
    $dayOffset = ([int]$Date.DayOfWeek + 6) % 7   # Montag = 0
    $thursday  = $Date.AddDays(3 - $dayOffset)
    return $thursday.Year
}

function Get-LastSunday {
    <#
    .SYNOPSIS
        Liefert das Datum des letzten Sonntags (relativ zu $From).
        Wenn $From selbst ein Sonntag ist, wird der vorherige Sonntag geliefert.
    #>
    param([datetime]$From = (Get-Date))

    $diff = ([int]$From.DayOfWeek)            # Sonntag = 0
    if ($diff -eq 0) { $diff = 7 }
    return $From.Date.AddDays(-$diff)
}

function Get-Bezeichnung {
    <#
    .SYNOPSIS
        Baut die Bezeichnung "Viafintech vom XX.KW YY".
    #>
    param(
        [datetime]$Date  = (Get-Date),
        [string]$Format  = 'Viafintech vom {0}.KW {1}'
    )

    $kw   = Get-IsoCalendarWeek -Date $Date
    $year = (Get-IsoYear -Date $Date).ToString().Substring(2,2)
    return ($Format -f $kw, $year)
}

function Get-OrdnerName {
    <#
    .SYNOPSIS
        Baut den Task-Ordnernamen, z.B. "26_06_02 Fuer 21.KW".
    .DESCRIPTION
        Das Datumspraefix ist der Lauftag ($Today), die Kalenderwoche stammt
        aus der zu bearbeitenden Woche ($RefDate, Stichtag = Sonntag). So
        entspricht der Ordnername exakt dem, was das Task-Skript erzeugt.
    #>
    param(
        [datetime]$RefDate = (Get-Date),
        [datetime]$Today   = (Get-Date),
        [string]$Format    = '{0} Fuer {1:D2}.KW'
    )

    $kw       = Get-IsoCalendarWeek -Date $RefDate
    $datePart = $Today.ToString('yy_MM_dd')
    return ($Format -f $datePart, $kw)
}

function Get-FilePrefix {
    <#
    .SYNOPSIS
        Baut das Datei-Praefix fuer den Zahllauf, z.B. "26_06_02_21_KW_".
    .DESCRIPTION
        Lauftag ($Today) plus Kalenderwoche der zu bearbeitenden Woche
        ($RefDate). Wird beim Umbenennen der Task-Dateien verwendet.
    #>
    param(
        [datetime]$RefDate = (Get-Date),
        [datetime]$Today   = (Get-Date)
    )

    $kw       = Get-IsoCalendarWeek -Date $RefDate
    $datePart = $Today.ToString('yy_MM_dd')
    return ('{0}_{1:D2}_KW_' -f $datePart, $kw)
}

function Get-ZahllaufContext {
    <#
    .SYNOPSIS
        Liefert ein PSCustomObject mit allen berechneten Werten.
    .DESCRIPTION
        Abgerechnet wird die VERGANGENE Woche: Der Stichtag ($RefDate) ist
        standardmaessig der letzte Sonntag. Seine ISO-Woche bestimmt KW, Jahr,
        Bezeichnung und OrdnerName; $RefDate ist zugleich die Faelligkeit.
        "Heute" (Now/Date/Weekday) bleibt das tatsaechliche Tagesdatum.
    .PARAMETER Today
        Tagesdatum fuer die "Heute"-Anzeige.
    .PARAMETER RefDate
        Stichtag (Sonntag) der zu bearbeitenden Woche. Default = letzter Sonntag.
    #>
    param(
        [datetime]$Today = (Get-Date),
        $RefDate         = $null
    )

    if (-not $RefDate) { $RefDate = Get-LastSunday -From $Today }
    $RefDate = [datetime]$RefDate

    [pscustomobject]@{
        Now           = $Today
        Date          = $Today.ToString('dd.MM.yyyy')
        Weekday       = $Today.ToString('dddd', [Globalization.CultureInfo]::GetCultureInfo('de-DE'))
        KW            = Get-IsoCalendarWeek -Date $RefDate
        Year          = Get-IsoYear -Date $RefDate
        Year2         = (Get-IsoYear -Date $RefDate).ToString().Substring(2,2)
        LastSunday    = $RefDate
        LastSundayStr = $RefDate.ToString('dd.MM.yyyy')
        Bezeichnung   = Get-Bezeichnung -Date $RefDate
        OrdnerName    = Get-OrdnerName -RefDate $RefDate -Today $Today
        FilePrefix    = Get-FilePrefix -RefDate $RefDate -Today $Today
    }
}

function Get-RecentWeeks {
    <#
    .SYNOPSIS
        Liefert die letzten N abrechenbaren Wochen (Stichtag = Sonntag),
        beginnend mit der vergangenen Woche. Jede Zeile traegt die KW und
        einen Anzeigetext "XX.KW YY mit Faelligkeit dd.MM.yyyy".
    #>
    param(
        [datetime]$From = (Get-Date),
        [int]$Count     = 4
    )

    $sunday = Get-LastSunday -From $From
    $list   = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $Count; $i++) {
        $d   = $sunday.AddDays(-7 * $i)
        $kw  = Get-IsoCalendarWeek -Date $d
        $yr2 = (Get-IsoYear -Date $d).ToString().Substring(2,2)
        $list.Add([pscustomobject]@{
            Sunday  = $d
            KW      = $kw
            Display = ('{0}.KW {1} mit Faelligkeit {2}' -f $kw, $yr2, $d.ToString('dd.MM.yyyy'))
        })
    }
    return $list
}
