$global:holidays = (Import-Csv .\holiday.csv ).date | ForEach-Object { get-date $_ }
function Test-BusinessDay {
    [CmdletBinding()]
    param (
        [datetime]$date
    )
    $valid = $true
    $weekend = ( ($date.DayOfWeek % 6) -eq 0 )
    $day = $date.Date
    $holiday = $global:holidays -contains $day
    $valid = !$weekend -and !$holiday
    return $valid
}
function Get-FirstBusinessDayBeforeDate {
    [CmdletBinding()]
    param (
        [datetime]$date,
        [int]$before = 0
    )
    $i = $before -1
    $valid = $false
    do {
        $i++
        $reminderDate = $date.AddDays(-$i)
        $valid = Test-BusinessDay $reminderDate
    } 
    while ( !$valid )
    [PSCustomObject]@{
        Date            = $date
        WeekDay         = $date.DayOfWeek
        BusinessDayBefore        = $reminderDate
        BusinessDayWeekDay = $reminderDate.DayOfWeek
        DateDiff = $i
    }
}