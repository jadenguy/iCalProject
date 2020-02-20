. $PSScriptRoot\CalendarEvent.ps1
. $PSScriptRoot\Get-BusinessDay.ps1
. $PSScriptRoot\ConvertTo-iCal.ps1
. $PSScriptRoot\Add-OutlookEvent.ps1
$global:holidays = (Import-Csv  $PSScriptRoot\holiday.csv ).date | ForEach-Object { get-date $_ }

# Gets eventEntries table
$resultsFolder = new-item -Path $PSScriptRoot -Name results -ItemType Directory -Force
$table = Get-Content  $PSScriptRoot\events.csv | ConvertFrom-Csv
# Creates list of each type of heading in eventEntries table
$calendars = $table | Group-Object -property 'TYPE'
# Goes through list again, once per event type
$calendars | ForEach-Object {
    $calendar = $_
    $calendarName = "$($calendar.Name) (Updated $((get-date).ToString('yyyy-MM-dd')))"
    # find relevant eventEntries in the table
    $eventEntries = @()
    $events = $calendar.Group
    $events | ForEach-Object {
        $event = $_
        $args = @{
            Summary     = $event.TYPE
            Description = Get-Content $PSScriptRoot\body.md -raw
            Location    = "707 W Lutz Lake Fern Rd, Lutz, FL 33548"
            Start       = (get-date $event.DATE).AddHours(7.5)
            End         = (get-date $event.DATE).AddHours(17)
        }
        $CalendarEvent = New-Object "CalendarEvent" -Property $args
        $CalendarEvent.SetReminder( $event.BEFORE, 0)
        if ($CalendarEvent.Validate()) {
            $eventEntries += $CalendarEvent
        }
    }
    # and this writes the file out
    $icalPath = join-path $resultsFolder "$calendarName.ics"
    $ical = $eventEntries | ConvertTo-iCal -calendar $calendarName 
    $ical | Set-Content $icalPath
    #".\results\$calendarName.ics"
    Start-Process $icalPath # uncomment this line to open each file on creation
    # $eventEntries | Add-OutlookEvent
    
}



# $args = @{
#     Summary     = "sum"
#     Description = "desc"
#     Start       = (get-date).AddHours(1)
#     End         = (get-date).AddHours(2)
#     Reminder = $True
#     ReminderDelta = "00:59:00"
# }
# $CalendarEvent = New-Object "CalendarEvent" -Property $args

# $CalendarEvent| ConvertTo-iCal|set-content test.ics
# Start-Process test.ics

# Useful links:
# partially using https://justinbraun.com/2018/01/powershell-dynamic-generation-of-an-ical-vcalendar-ics-format-file/
# https://stackoverflow.com/questions/35645402/how-to-specify-timezone-in-ics-file-which-will-work-efficiently-with-google-outl
# https://apps.marudot.com/ical/
# AND TO IMPORT https://thescriptkeeper.wordpress.com/2013/09/27/import-a-bunch-of-ics-calendar-files-with-powershell/
# We then use this knowledge to just inject into outlook

# SharePoint Shared Calendar (another direciotn we can take this project)
# stssync://sts/?ver=1.1&type=calendar&cmd=add-folder&base-url=https%3A%2F%2Fattentiem%2Esharepoint%2Ecom%2Fsites%2Fusac%2Emmm&list-url=%2FLists%2FBiWeekly%2520Payroll%2F&guid=%7B9cd4bb9e%2Df405%2D40f8%2D9727%2D614ea89a15b1%7D&site-name=Attenti%20Electronic%20Monitoring&list-name=Bi%2DWeekly%20Payroll