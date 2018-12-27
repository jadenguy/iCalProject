class CalendarEvent {
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
    CalendarEvent() {
        $this.Summary = "New Event"
        $this.Description = $this.Summary
        $this.Start = Get-Date
        $this.End = $this.Start.addhours(1)
    }
    CalendarEvent([String] $EventName) {
        $this.Summary = $EventName
        $this.Description = $this.Summary
        $this.Start = Get-Date
        $this.End = $this.Start.addhours(1)
    }
    CalendarEvent([Datetime] $StartDate) {
        $this.Summary = "New Event"
        $this.Description = $this.Summary
        $this.Start = $StartDate
        $this.End = $this.Start.addhours(1)
    }
    CalendarEvent([String] $EventName, [Datetime] $StartDate) {
        $this.Summary = $EventName
        $this.Description = $this.Summary
        $this.Start = $StartDate
        $this.End = $this.Start.addhours(1)
    }
    CalendarEvent([String] $EventName, [datetime] $StartDate, [int] $BusinessDaysBefore, [int] $HoursBefore = 0) {
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
    [CalendarEvent]static Create() {
        $event = New-Object "CalendarEvent"
        return $event
    }
}