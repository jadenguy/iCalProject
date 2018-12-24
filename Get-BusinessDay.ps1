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
    # since we're adding 1 to the day before calculating if it's a business day every time, even if you want the same date, first subtract a day
    $i = $before - 1
    # run through the logic once and get an answer for valid
    do {
        $i++
        # checks the reminder date in question
        $reminderDate = $date.AddDays(-$i)
        $valid = Test-BusinessDay $reminderDate
    } 
    while ( !$valid )
    # returns an object containing the date you asked for, it's weekday, the reminder date, it's weekday, and how many days back that is, useful since we multiply that to hours for the reminder date
    [PSCustomObject]@{
        Date               = $date
        WeekDay            = $date.DayOfWeek
        BusinessDayBefore  = $reminderDate
        BusinessDayWeekDay = $reminderDate.DayOfWeek
        DateDiff           = $i
    }
}