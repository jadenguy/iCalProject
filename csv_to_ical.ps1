$csvContent = Get-Content .\data.csv
$table = $csvContent|ConvertFrom-Csv
$calendars = $table.TYPE|Sort-Object -Unique
$calendars | ForEach-Object {
    $calendar = $_
    
    $events = $table|Where-Object -Property TYPE -eq $calendar
    # $events | ConvertTo-Csv | Set-Content "$calendar.csv"
    "$calendar.ics"
    $events |Format-Table
}


# CONSIDERING USING https://justinbraun.com/2018/01/powershell-dynamic-generation-of-an-ical-vcalendar-ics-format-file/