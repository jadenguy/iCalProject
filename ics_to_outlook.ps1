# Converts single ICS file to 
$ICSpath = ".\results\"
$ICSlist = get-childitem $ICSPath -Filter sample.ics
$ICSlist 

Foreach ($i in $ICSlist ) {
    $file = $i. fullname

    $content = Get-Content $file -Encoding UTF8
    $content
    $file
    $data = @{}
    $content | foreach-Object {
        $line = $_
        try {
            $key = ($line.split( ':').split( ';')[0])
            $value = ( $line.split( ':')[1]).Trim()
            if ($key) {
                $data[$key] = $value
            }
        }
        catch {
        }   
    }

    $Body = [regex]::match($content, '(?<=\DESCRIPTION:).+(?=\DTEND:)', "singleline").value.trim()
    $Body = $Body -replace "\r\n\s"
    $Body = $Body.replace("\,", ",").replace("\n", " ")
    $Body = $Body -replace "\s\s"

    $Start = ($data.getEnumerator() | ? { $_.Name -eq "DTSTART"}).Value -replace "T"
    $Start = [datetime]::ParseExact($Start , "yyyyMMddHHmmss" , $null )

    $End = ($data.getEnumerator() | ? { $_.Name -eq "DTEND"}).Value -replace "T"
    $End = [datetime]::ParseExact($End , "yyyyMMddHHmmss" , $null )

    $Subject = ($data.getEnumerator() | ? { $_.Name -eq "SUMMARY"}).Value
    $Location = ($data.getEnumerator() | ? { $_.Name -eq "LOCATION"}).Value
    $eventObject = [PSCustomObject]@{
        Body     = $Body
        Start    = $Start
        End      = $End
        Subject  = $Subject
        Location = $Location
    }
    $outlook = new-object -com Outlook.Application
    $calendar = $outlook.Session.GetDefaultFolder(9)
    $appt = $calendar.Items.Add(1)

    $appt.Start = $Start
    $appt.End = $End
    $appt.Subject = $Subject
    $appt.Categories = "Presentations" #Pick your own category!
    $appt.BusyStatus = 0   # 0=Free
    $appt.Location = $Location
    $appt.Body = $Body
    $appt.ReminderMinutesBeforeStart = 15 #Customize if you want 
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