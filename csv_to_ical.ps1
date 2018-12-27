. $PSScriptRoot\ConvertTo-iCal.ps1
. $PSScriptRoot\Add-OutlookEvent.ps1

# Gets eventEntries table
new-item -Path $PSScriptRoot -Name results -ItemType Directory -Force | Out-Null
$table = Get-Content .\events.csv|ConvertFrom-Csv
# Creates list of each type of heading in eventEntries table
$calendars = $table | Group-Object -property 'TYPE'
# Goes through list again, once per event type
$calendars | ForEach-Object {
    $calendar = $_
    $calendarName = $calendar.Name
    # find relevant eventEntries in the table
    $eventEntries = @()
    $events = $calendar.Group
    $events | ForEach-Object {
        $event = $_
        $args = @{
            Summary     = $event.TYPE
            Description = $event.TYPE
            Start       = (get-date $event.DATE).AddHours(8)
            End         = (get-date $event.DATE).AddHours(12)
        }
        $CalendarEvent = New-Object "CalendarEvent" -Property $args
        $CalendarEvent.SetReminder( $event.BEFORE, 0)
        if ($CalendarEvent.Validate()) {
            $eventEntries += $CalendarEvent
        }
    }
    # and this writes the file out
    # $eventEntries | ConvertTo-iCal -calendar $calendarName|Set-Content ".\results\$calendarName.ics"
    # Start-Process ".\results\$calendarName.ics" # uncomment this line to open each file on creation
    # $eventEntries | Add-OutlookEvent
    
}

# Useful links:
# partially using https://justinbraun.com/2018/01/powershell-dynamic-generation-of-an-ical-vcalendar-ics-format-file/
# https://stackoverflow.com/questions/35645402/how-to-specify-timezone-in-ics-file-which-will-work-efficiently-with-google-outl
# https://apps.marudot.com/ical/
# AND TO IMPORT https://thescriptkeeper.wordpress.com/2013/09/27/import-a-bunch-of-ics-calendar-files-with-powershell/
# But at this point we can just create new library items based on the above criteria, and we can skip the ICS stuff really

# SharePoint Shared Calendar (another direciotn we can take this project)
# stssync://sts/?ver=1.1&type=calendar&cmd=add-folder&base-url=https%3A%2F%2Fattentiem%2Esharepoint%2Ecom%2Fsites%2Fusac%2Emmm&list-url=%2FLists%2FBiWeekly%2520Payroll%2F&guid=%7B9cd4bb9e%2Df405%2D40f8%2D9727%2D614ea89a15b1%7D&site-name=Attenti%20Electronic%20Monitoring&list-name=Bi%2DWeekly%20Payroll