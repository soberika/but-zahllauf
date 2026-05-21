# =============================================================================
#  DateHelpers.ps1
#  Datumsfunktionen fuer den viafintech Zahllauf
# =============================================================================

function Get-IsoCalendarWeek {
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
        Baut den Task-Ordnernamen, z.B. "26_05_21 Fuer 20.KW".
    #>
    param(
        [datetime]$Date = (Get-Date),
        [string]$Format = '{0} Fuer {1}.KW'
    )

    $kw      = Get-IsoCalendarWeek -Date $Date
    $datePart = $Date.ToString('yy_MM_dd')
    return ($Format -f $datePart, $kw)
}

function Get-ZahllaufContext {
    <#
    .SYNOPSIS
        Liefert ein PSCustomObject mit allen aktuell berechneten Werten.
    #>
    param([datetime]$Date = (Get-Date))

    [pscustomobject]@{
        Now           = $Date
        Date          = $Date.ToString('dd.MM.yyyy')
        Weekday       = $Date.ToString('dddd', [Globalization.CultureInfo]::GetCultureInfo('de-DE'))
        KW            = Get-IsoCalendarWeek -Date $Date
        Year          = Get-IsoYear -Date $Date
        Year2         = (Get-IsoYear -Date $Date).ToString().Substring(2,2)
        LastSunday    = Get-LastSunday -From $Date
        LastSundayStr = (Get-LastSunday -From $Date).ToString('dd.MM.yyyy')
        Bezeichnung   = Get-Bezeichnung -Date $Date
        OrdnerName    = Get-OrdnerName -Date $Date
    }
}
