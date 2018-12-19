$csvContent = Get-Content .\data.csv
$table = $csvContent|ConvertFrom-Csv
$table