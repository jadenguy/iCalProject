. $PSScriptRoot\Get-BusinessDay.ps1
$longUTCDateFormat = "yyyyMMddTHHmmssZ"
$longDateFormat = "yyyyMMddTHHmmss"
$table = Get-Content .\events.csv|ConvertFrom-Csv
$calendars = $table.TYPE|Sort-Object -Unique
$calendars | ForEach-Object {
    $calendar = $_
    $ical = [System.Text.StringBuilder]::new()
    [void]$ical.AppendLine('BEGIN:VCALENDAR')
    [void]$ical.AppendLine('VERSION:2.0')
    [void]$ical.AppendLine('METHOD:PUBLISH')
    [void]$ical.AppendLine('PRODID:-//Braunweb//PowerShell ICS Creator Sample//EN')
    [void]$ical.AppendLine('X-WR-CALNAME:' + $calendar)
    
    $events = $table|Where-Object -Property TYPE -eq $calendar
    $events | ForEach-Object {
        $event = $_
        $reminder = Get-FirstBusinessDayBeforeDate -date (get-date $event.date).addhours(8) -before $event.before
        [void]$ical.AppendLine('BEGIN:VEVENT')
        [void]$ical.AppendLine("UID:" + [guid]::NewGuid())
        [void]$ical.AppendLine("CREATED:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
        [void]$ical.AppendLine("DTSTAMP:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
        [void]$ical.AppendLine("LAST-MODIFIED:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
        [void]$ical.AppendLine("SEQUENCE:0")
        [void]$ical.AppendLine("DTSTART:" + $reminder.date.ToString($longDateFormat))
        [void]$ical.AppendLine("DTEND:" + $reminder.date.addhours(4).ToString($longDateFormat))
        [void]$ical.AppendLine("DESCRIPTION:" + $calendar)
        [void]$ical.AppendLine("SUMMARY:" + $calendar)
        [void]$ical.AppendLine("LOCATION:" + $calendar)
        [void]$ical.AppendLine("TRANSP:TRANSPARENT")        
        [void]$ical.AppendLine("BEGIN:VALARM")
        [void]$ical.AppendLine("ACTION:DISPLAY")
        [void]$ical.AppendLine("DESCRIPTION:Submit $calendar")
        [void]$ical.AppendLine("TRIGGER:-P0DT$($reminder.DateDiff*24)H0M0S")
        [void]$ical.AppendLine("END:VALARM")
        [void]$ical.AppendLine('END:VEVENT')
    }
    [void]$ical.AppendLine('END:VCALENDAR')
    Set-Content -Value $ical -path (join-path ".\results" "$calendar.ics")
}


# partially using https://justinbraun.com/2018/01/powershell-dynamic-generation-of-an-ical-vcalendar-ics-format-file/
# stssync://sts/?ver=1.1&type=calendar&cmd=add-folder&base-url=https%3A%2F%2Fattentiem%2Esharepoint%2Ecom%2Fsites%2Fusac%2Emmm&list-url=%2FLists%2FBiWeekly%2520Payroll%2F&guid=%7B9cd4bb9e%2Df405%2D40f8%2D9727%2D614ea89a15b1%7D&site-name=Attenti%20Electronic%20Monitoring&list-name=Bi%2DWeekly%20Payroll