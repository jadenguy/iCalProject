$global:holidays = (Import-Csv .\holiday.csv ).date | ForEach-Object { get-date $_ }
#this function checks to see if a day is a weekend or a holiday, and returns a true (m-f) or false (s-s)
function Test-BusinessDay {
    [CmdletBinding()]
    param (
        [datetime]$date
    )
    # weekends include days 0 and 6
    $weekend = ( (0, 6) -contains $date.DayOfWeek  )
    # midnights considered only from global holiday list
    $holiday = $global:holidays -contains $date.Date
    # it's only a business day if it's not a weekend or holiday, so return true then
    $valid = !$weekend -and !$holiday
    return $valid
}
function Get-FirstBusinessDayBeforeDate {
    [CmdletBinding()]
    param (
        [datetime]$date = (get-date),
        [int]$before = 0
    )
    do {
        do {
            $businessDaysBack = $i + $skipDays
            $testDate = $date.AddDays(-$businessDaysBack)
            $passed = Test-BusinessDay $testDate 
            if (!$passed) {
                $skipDays++
            }
        } while (!$passed)
        $i++
    } while ($i -le $before)
    # uses that date
    # returns an object containing the date you asked for, it's weekday, the reminder date, it's weekday, and how many days back that is, useful since we multiply that to hours for the reminder date
    [PSCustomObject]@{
        Date               = $date
        WeekDay            = $date.DayOfWeek
        BusinessDayBefore  = $testDate
        BusinessDayWeekDay = $testDate.DayOfWeek
        DateDiff           = $businessDaysBack
        BusinessDateDiff   = $before
    }
}