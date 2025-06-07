# Export-FastExcel
PowershellでEXCELファイルを出力するサンプル
## 概要
[RubyでEXCELファイルを出力するサンプル](https://github.com/gearsns/RubyFastExcel)でruby版を作ったので同じようにPowershell版も実装してみた

## 使い方

import
```` powershell
. "$PSScriptRoot\Export-FastExcel.ps1"
````

例1 引数
```` powershell
$data = @(
    @("A", "B""", "A"),
    @("1", 2, [Datetime]"2019/12/31 01:23:45")
)

Export-FastExcel -Filename "./Output/Output1.xlsx" -Data $data
````

例2 パイプで
```` powershell
$data = @(
    @("A", "B""", "A"),
    @("1", 2, [Datetime]"2019/12/31 01:23:45")
)

$data | Export-FastExcel -Filename "./Output/Output2.xlsx"
````

例3 プロセス一覧
```` powershell
Get-Process | Select-Object Name, Id, WorkingSet | Export-FastExcel -Filename "./Output/Processes.xlsx"
````

例4 オブジェクト
```` powershell
"Hello", "World" | ForEach-Object { [PSCustomObject]@{ Message = $_; Length = $_.Length } } | Export-FastExcel -Filename "./Output/Messages.xlsx"
````