. $PSScriptRoot\Get-BusinessDay.ps1
class IcsEvent {
    # Properties
    [String] $Description
    [String] $Summary
    [String] $Location
    [Boolean]$Busy
    [Datetime] $Start
    [Datetime] $End
    [TimeSpan] $ReminderDelta
    [Boolean] $Reminder = $false

    # Constructors
    IcsEvent() {
        $this.Summary = "New Event"
        $this.Description = $this.Summary
        $this.Start = Get-Date
        $this.End = $this.Start.addhours(1)
    }
    IcsEvent([String] $EventName) {
        $this.Summary = $EventName
        $this.Description = $this.Summary
        $this.Start = Get-Date
        $this.End = $this.Start.addhours(1)
    }
    IcsEvent([Datetime] $StartDate) {
        $this.Summary = "New Event"
        $this.Description = $this.Summary
        $this.Start = $StartDate
        $this.End = $this.Start.addhours(1)
    }
    IcsEvent([String] $EventName, [Datetime] $StartDate) {
        $this.Summary = $EventName
        $this.Description = $this.Summary
        $this.Start = $StartDate
        $this.End = $this.Start.addhours(1)
    }
    IcsEvent([String] $EventName, [datetime] $StartDate, [int] $BusinessDaysBefore, [int] $HoursBefore = 0) {
        $this.Summary = $EventName
        $this.Description = $this.Summary
        $this.Start = $StartDate
        $this.End = $this.Start.addhours(1)
        $this.ReminderDelta = New-Timespan -Days (Get-FirstBusinessDayBeforeDate -date ($StartDate) -before $BusinessDaysBefore).DateDiff -Hours $HoursBefore
        $this.Reminder = $true
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
    [Boolean] Validate() {
        $starts = ($this.Start -ne 0)
        $startsBeforeEnds = ($this.Start -le $this.End)
        $reminderNotAllowed = !($this.Reminder)
        if ($reminderNotAllowed) {
            $reminderValid = ($this.ReminderDelta.TotalMilliseconds -eq 0)
        }
        else {
            $reminderValid = $true
        }

        # write-host $this.ToString()
        return $starts -and $startsBeforeEnds -and $reminderValid
    }
    [void] SetReminder([int]$BusinessDaysBefore, [int] $HoursBefore = 0) {
        $this.Reminder = $true
        $this.ReminderDelta =  New-Timespan -Days (Get-FirstBusinessDayBeforeDate -date ($this.Start) -before $BusinessDaysBefore).DateDiff -Hours $HoursBefore
    }
    [void] SetReminder([int]$BusinessDaysBefore) {
        $this.Reminder = $true
        $this.ReminderDelta =  New-Timespan -Days (Get-FirstBusinessDayBeforeDate -date ($this.Start) -before $BusinessDaysBefore).DateDiff
    }
    [IcsEvent]static Create() {
        $event = New-Object "IcsEvent"
        return $event
    }
}
function ConvertTo-iCal {
    [CmdletBinding()]
    param (
        [string]$calendar = "New_Calendar",
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)][IcsEvent]$event,
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
        [void]$ical.AppendLine('PRODID:New_Ics_PowerShell_Script')
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
        if ($event.Validate()) {
            [void]$ical.AppendLine('BEGIN:VEVENT')
            [void]$ical.AppendLine("UID:" + [guid]::NewGuid())
            [void]$ical.AppendLine("CREATED:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
            [void]$ical.AppendLine("DTSTAMP:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
            [void]$ical.AppendLine("LAST-MODIFIED:" + [datetime]::Now.ToUniversalTime().ToString($longUTCDateFormat))
            [void]$ical.AppendLine("SEQUENCE:0")
            [void]$ical.AppendLine("DTSTART;TZID=America/New_York:" + $event.Start.ToString($longDateFormat))
            [void]$ical.AppendLine("DTEND;TZID=America/New_York:" + $event.End.ToString($longDateFormat))
            [void]$ical.AppendLine("DESCRIPTION:" + $event.Description)
            [void]$ical.AppendLine("SUMMARY:" + $event.Summary)
            [void]$ical.AppendLine("LOCATION:" + $event.Location)            
            [void]$ical.AppendLine("TRANSP:$(if ($event.Busy){"OPAQUE"} else {"TRANSPARENT"})")
            if ($event.Reminder) {
                [void]$ical.AppendLine("BEGIN:VALARM")
                [void]$ical.AppendLine("ACTION:DISPLAY")
                [void]$ical.AppendLine("DESCRIPTION:Reminder: " + $event.Summary)
                [void]$ical.AppendLine("TRIGGER:-P$($event.ReminderDelta.Days)DT$($event.ReminderDelta.Hours)H$($event.ReminderDelta.Minutes)M$($event.ReminderDelta.Seconds)S")
                [void]$ical.AppendLine("END:VALARM")
            }
            [void]$ical.AppendLine('END:VEVENT')
        }
    }
    end {
        [void]$ical.AppendLine('END:VCALENDAR')
        Write-Output $ical.ToString()
    }
}

# $x = [IcsEvent]::new('hello', '2019-01-01 12:00')
# $y = [IcsEvent]::new('hello', '2019-01-02 12:00')
# $z = [IcsEvent]::new('hello', '2019-01-03 12:00', 0, 4)
# $x.Validate()
# $y.validate()
# $z.Validate()
# $x.Busy = $true
# $ical = $x, $y, $z|ConvertTo-iCal -calendar "hello_calendar" 
# $ical | Set-Content hello.ics; Start-Process hello.ics