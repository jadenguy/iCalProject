$start = get-date "3/3/2020"

$days = 0..12 | ForEach-Object { $week = $_ 
    0..5|ForEach-Object{
        $day = $_
        $start.AddDays($day).AddDays($week *7).ToString("%yyy-MM-dd")
    }
}
$days | set-clipboard;get-clipboard