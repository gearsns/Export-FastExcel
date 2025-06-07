. "$PSScriptRoot\Export-FastExcel.ps1"

$data = @(
    @("A", "B""", "A"),
    @("1", 2, [Datetime]"2019/12/31 01:23:45")
)

Export-FastExcel -Filename "./Output/Output1.xlsx" -Data $data
$data | Export-FastExcel -Filename "./Output/Output2.xlsx"

Get-Process | Select-Object Name, Id, WorkingSet | Export-FastExcel -Filename "./Output/Processes.xlsx"

"Hello", "World" | ForEach-Object { [PSCustomObject]@{ Message = $_; Length = $_.Length } } | Export-FastExcel -Filename "./Output/Messages.xlsx"
