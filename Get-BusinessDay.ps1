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
        [datetime]$date,
        [int]$before = 0
    )
    # start with i days back as guessed by the 'before' date, then subtract days from there
    $i = $before
    # run through the logic once and get an answer for valid, and keeps looking back futher until it finds it.
    while ( !( Test-BusinessDay $date.AddDays(-$i) ) ) {
        $i++
    } 
    # uses that date
    $reminderDate = $date.AddDays(-$i)
    # returns an object containing the date you asked for, it's weekday, the reminder date, it's weekday, and how many days back that is, useful since we multiply that to hours for the reminder date
    [PSCustomObject]@{
        Date               = $date
        WeekDay            = $date.DayOfWeek
        BusinessDayBefore  = $reminderDate
        BusinessDayWeekDay = $reminderDate.DayOfWeek
        DateDiff           = $i
    }
}