function Add-OutlookEvent {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)][CalendarEvent]$event
    )

    
    begin {
        $outlook = new-object -com Outlook.Application
        $calendar = $outlook.Session.GetDefaultFolder(9) 
    }
    
    process {

        if ($event.Validate()) {  
            $appt = $calendar.Items.Add(1)
            $appt.Start = $event.Start
            $appt.End = $event.End
            $appt.Subject = $event.Summary
            $appt.Categories = "Presentations" #Pick your own category!
            $appt.BusyStatus = 0   # 0=Free
            $appt.Location = $event.Location
            $appt.Body = $event.Description
            if ($event.Reminder)
            {$appt.ReminderMinutesBeforeStart = $event.ReminderDelta.TotalMinutes}
            $appt.Save()
            $apptObject = [PSCustomObject]@{
                Start                      = $appt.Start
                End                        = $appt.End
                Subject                    = $appt.Subject
                Categories                 = $appt.Categories
                BusyStatus                 = $appt.BusyStatus
                Location                   = $appt.Location
                Body                       = $appt.Body
                ReminderMinutesBeforeStart = $appt.ReminderMinutesBeforeStart
                Saved                      = $appt.Saved
            }
            if ($appt.Saved)
            { write-host "Appointment saved."}
            Else {write-host "Appointment NOT saved."}
            return $apptObject
        }
    }
    
    
    end {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Write-Debug
        $Outlook = $null
    }
}


# $newEvent = New-Object CalendarEvent
# $newEvent.Start = (Get-Date).AddDays(1)
# $newEvent.End = $newEvent.Start.AddHours(1)
# $newEvent.Validate()
# write-output $newEvent
# Add-OutlookEvent $newEvent|ft -AutoSize