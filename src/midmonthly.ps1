. $PSScriptRoot\CalendarEvent.ps1
. $PSScriptRoot\Get-BusinessDay.ps1
. $PSScriptRoot\ConvertTo-iCal.ps1
. $PSScriptRoot\Add-OutlookEvent.ps1
$global:holidays = (Import-Csv  $PSScriptRoot\holiday.csv ).date | ForEach-Object { get-date $_ }

# Gets eventEntries table
new-item -Path $PSScriptRoot -Name results -ItemType Directory -Force | Out-Null
$table = Get-Content $PSScriptRoot\events.csv|ConvertFrom-Csv
# Creates list of each type of heading in eventEntries table
$wanted = "Mid-Month Expense Reimbursements"
$calendar = $table | Where-Object -Property "TYPE" -EQ -Value $wanted
$calendarName = "$($wanted) (Updated $((get-date).ToString('yyyy-MM-dd')))"
# find relevant eventEntries in the table
$eventEntries = @()
$events = $calendar
$events | ForEach-Object {
    $event = $_
    $args = @{
        Summary     = $event.TYPE
        Description = $calendarName
        Start       = (get-date $event.DATE).AddHours(8)
        End         = (get-date $event.DATE).AddHours(12)
    }
    $CalendarEvent = New-Object "CalendarEvent" -Property $args
    $CalendarEvent.SetReminder( $event.BEFORE, 0)
    if ($CalendarEvent.Validate()) {
        $eventEntries += $CalendarEvent
    }
}
$eventEntries | Add-OutlookEvent