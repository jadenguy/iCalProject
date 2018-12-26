. $PSScriptRoot\Get-BusinessDay.ps1
$longUTCDateFormat = "yyyyMMddTHHmmssZ"
$longDateFormat = "yyyyMMddTHHmmss"
# Gets events table
new-item -Path $PSScriptRoot -Name results -ItemType Directory -Force | Out-Null
$table = Get-Content .\events.csv|ConvertFrom-Csv
# Creates list of each type of heading in events table
$calendars = $table.TYPE|Sort-Object -Unique
# Goes through list again, once per event type
$calendars | ForEach-Object {
    $calendar = $_
    # Starts string builder, one way to create lines. This is easier and harder to read than using a raw string tool, but you don't have to mess around with making your own newline characters
    $ical = [System.Text.StringBuilder]::new()
    [void]$ical.AppendLine('BEGIN:VCALENDAR')
    [void]$ical.AppendLine('VERSION:2.0')
    [void]$ical.AppendLine('METHOD:PUBLISH')
    [void]$ical.AppendLine('PRODID:Alfredo_PowerShell_Script')
    [void]$ical.AppendLine('X-WR-CALNAME:' + $calendar) # This is where the title in Outlook comes from. Took a while to get this one right.
    [void]$ical.AppendLine("X-WR-TIMEZONE:America/New_York")
    [void]$ical.AppendLine("BEGIN:VTIMEZONE")
    [void]$ical.AppendLine("TZID:America/New_York")
    [void]$ical.AppendLine("X-LIC-LOCATION:America/New_York")
    [void]$ical.AppendLine("BEGIN:DAYLIGHT")
    [void]$ical.AppendLine("TZOFFSETFROM:-0500")
    [void]$ical.AppendLine("TZOFFSETTO:-0400")
    [void]$ical.AppendLine("TZNAME:EDT")
    [void]$ical.AppendLine("DTSTART:19700308T020000")
    [void]$ical.AppendLine("RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU")
    [void]$ical.AppendLine("END:DAYLIGHT")
    [void]$ical.AppendLine("BEGIN:STANDARD")
    [void]$ical.AppendLine("TZOFFSETFROM:-0400")
    [void]$ical.AppendLine("TZOFFSETTO:-0500")
    [void]$ical.AppendLine("TZNAME:EST")
    [void]$ical.AppendLine("DTSTART:19701101T020000")
    [void]$ical.AppendLine("RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU")
    [void]$ical.AppendLine("END:STANDARD")
    [void]$ical.AppendLine("END:VTIMEZONE")
    # find relevant events in the table
    $events = $table|Where-Object -Property TYPE -eq $calendar
    $events | ForEach-Object {
        $event = $_
        # this is a custom funciton i wrote to find the last business day before a given date, x days before. I made the events start at 8 AM.
        $reminder = Get-FirstBusinessDayBeforeDate -date (get-date $event.date).addhours(8) -before $event.before
        # this is just writing an entry in the format tht ICS files requires, mostly taken from the first link
        [void]$ical.AppendLine('BEGIN:VEVENT')
        [void]$ical.AppendLine("UID:" + [guid]::NewGuid())
        [void]$ical.AppendLine("CREATED:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
        [void]$ical.AppendLine("DTSTAMP:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
        
        [void]$ical.AppendLine("LAST-MODIFIED:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
        [void]$ical.AppendLine("SEQUENCE:0")
        [void]$ical.AppendLine("DTSTART;TZID=America/New_York:" + $reminder.date.ToString($longDateFormat))
        [void]$ical.AppendLine("DTEND;TZID=America/New_York:" + $reminder.date.addhours(4).ToString($longDateFormat))
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
    # this is how to close an ics file
    [void]$ical.AppendLine('END:VCALENDAR')
    # and this writes the file out
    Set-Content -Value $ical -path (join-path ".\results" "$calendar.ics")
}

# Useful links:
# partially using https://justinbraun.com/2018/01/powershell-dynamic-generation-of-an-ical-vcalendar-ics-format-file/
# https://stackoverflow.com/questions/35645402/how-to-specify-timezone-in-ics-file-which-will-work-efficiently-with-google-outl
# https://apps.marudot.com/ical/
# SharePoint Shared Calendar (another direciotn we can take this project)
# stssync://sts/?ver=1.1&type=calendar&cmd=add-folder&base-url=https%3A%2F%2Fattentiem%2Esharepoint%2Ecom%2Fsites%2Fusac%2Emmm&list-url=%2FLists%2FBiWeekly%2520Payroll%2F&guid=%7B9cd4bb9e%2Df405%2D40f8%2D9727%2D614ea89a15b1%7D&site-name=Attenti%20Electronic%20Monitoring&list-name=Bi%2DWeekly%20Payroll