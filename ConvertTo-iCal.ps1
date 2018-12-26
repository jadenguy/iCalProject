. $PSScriptRoot\Get-BusinessDay.ps1
class IcsEvent {
    # Properties
    [String] $Description
    [String] $Summary
    [String] $Location
    [Datetime] $Start
    [Datetime] $End
    [TimeSpan] $ReminderDelta

    # Constructors
    IcsEvent([String] $EventName) {
        $this.Summary = $EventName
        $this.Start = Get-Date
        $this.ReminderDelta = New-TimeSpan
    }
    IcsEvent([Datetime] $StartDate) {
        $this.Summary = "New Event"
        $this.Start = $StartDate
        $this.ReminderDelta = New-TimeSpan
    }
    IcsEvent([String] $EventName, [Datetime] $StartDate) {
        $this.Summary = $EventName
        $this.Start = $StartDate
        $this.ReminderDelta = New-TimeSpan
    }
    IcsEvent([String] $EventName, [datetime] $StartDate, [int] $BusinessDaysBefore, [int] $HoursBefore = 0) {
        $this.Summary = $EventName
        $this.Start = $StartDate
        $this.ReminderDelta = New-Timespan -Days (Get-FirstBusinessDayBeforeDate -date ($StartDate) -before $BusinessDaysBefore).DateDiff -Hours $HoursBefore
    }

    # Methods
    [string] ToString() { return $this.Render() }
    hidden [string] Render() {
        $ret = $this.Summary 
        $ret += " starts at "
        $ret += $this.Start
        if ($this.End -ne 0) { 
            $ret += " ends at "
            $ret += $this.End
            $ret += " lasting "
            $ret += New-TimeSpan -Start $this.Start -End $this.End
        }
        $ret += " and will alarm "
        $ret += $this.ReminderDelta.ToString()
        $ret += " before."        
        return $ret
    }
}
function ConvertTo-iCal {
    [CmdletBinding()]
    param (
        [string]$calendar = "New_Calendar",
        [IcsEvent]$event,
        [TimeZoneInfo]$tz
    )

    begin {
        if (!$tz) {
            $tz = Get-TimeZone
        }
        $longUTCDateFormat = "yyyyMMddTHHmmssZ"
        $longDateFormat = "yyyyMMddTHHmmss"
        $ical = [System.Text.StringBuilder]::new()
        [void]$ical.AppendLine('BEGIN:VCALENDAR')
        [void]$ical.AppendLine('VERSION:2.0')
        [void]$ical.AppendLine('METHOD:PUBLISH')
        [void]$ical.AppendLine('PRODID:Alfredo_PowerShell_Script')
        [void]$ical.AppendLine('X-WR-CALNAME:' + $calendar)
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
    }

    process {            
        # this is just writing an entry in the format tht ICS files requires, mostly taken from the first link
        if ($event) {
            [void]$ical.AppendLine('BEGIN:VEVENT')
            [void]$ical.AppendLine("UID:" + [guid]::NewGuid())
            [void]$ical.AppendLine("CREATED:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
            [void]$ical.AppendLine("DTSTAMP:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
            [void]$ical.AppendLine("LAST-MODIFIED:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
            [void]$ical.AppendLine("SEQUENCE:0")
            [void]$ical.AppendLine("DTSTART;TZID=America/New_York:" + $event.Start.ToString($longDateFormat))
            [void]$ical.AppendLine("DTEND;TZID=America/New_York:" + $event.End.ToString($longDateFormat))
            [void]$ical.AppendLine("DESCRIPTION:" + $calendar)
            [void]$ical.AppendLine("SUMMARY:" + $calendar)
            [void]$ical.AppendLine("LOCATION:" + $calendar)
            [void]$ical.AppendLine("TRANSP:TRANSPARENT")        
            if ($event.ReminderDelta.TotalMilliseconds) {
                [void]$ical.AppendLine("BEGIN:VALARM")
                [void]$ical.AppendLine("ACTION:DISPLAY")
                [void]$ical.AppendLine("DESCRIPTION:Submit $calendar")
                #this is where we use the reminder days before        
                [void]$ical.AppendLine("TRIGGER:-P$($event.ReminderDelta.Days)DT$($event.ReminderDelta.Hours)H$($event.ReminderDelta.Minutes)M$($event.ReminderDelta.Seconds)S")
                [void]$ical.AppendLine("END:VALARM")
                [void]$ical.AppendLine('END:VEVENT')
            }
        }
    }

    end {
        [void]$ical.AppendLine('END:VCALENDAR')
        Write-Output $ical.ToString()
    }
}

$x = [IcsEvent]::new('hello', '2018-01-01 12:00', 1, 4)
$x.End = get-date
ConvertTo-iCal | Out-GridView
ConvertTo-iCal -event $x -calendar "hello_calendar"