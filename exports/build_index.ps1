$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $PSCommandPath

function Get-Csv {
  param([string]$name)
  $path = Join-Path $dir $name
  if (Test-Path $path) { Import-Csv -Path $path } else { @() }
}

function HtmlEncode([string]$s) { [System.Net.WebUtility]::HtmlEncode($s) }

function Render {
  param($rows,[string]$title)
  if ($rows.Count -eq 0) { return "<section><h2>$title</h2><p>No data</p></section>" }
  $cols = $rows[0].PSObject.Properties.Name
  $sb = "<section><h2>$title</h2><table><thead><tr>"
  foreach ($c in $cols) { $sb += "<th>" + (HtmlEncode $c) + "</th>" }
  $sb += "</tr></thead><tbody>"
  foreach ($r in $rows) {
    $sb += "<tr>"
    foreach ($c in $cols) { $sb += "<td>" + (HtmlEncode ($r.$c)) + "</td>" }
    $sb += "</tr>"
  }
  $sb += "</tbody></table></section>"
  return $sb
}

$sections = @()
$sections += (Render (Get-Csv 'global_finance_totals.csv') '全量合计')
$sections += (Render (Get-Csv 'project_finance_summary.csv') '项目汇总明细')
$sections += (Render (Get-Csv 'monthly_contracts.csv') '月签约')
$sections += (Render (Get-Csv 'monthly_payments.csv') '月回款')
$sections += (Render (Get-Csv 'project_completion_rank.csv') '完成率排行')
$sections += (Render (Get-Csv 'quarterly_contracts.csv') '季度签约')
$sections += (Render (Get-Csv 'quarterly_payments.csv') '季度回款')
$sections += (Render (Get-Csv 'yearly_contracts.csv') '年度签约')
$sections += (Render (Get-Csv 'yearly_payments.csv') '年度回款')
$sections += (Render (Get-Csv 'overdue_projects_90d.csv') '逾期未回款>90天')

$style = 'body{font-family:Segoe UI,Arial;margin:20px} h1{margin-top:0} table{border-collapse:collapse;width:100%} th,td{border:1px solid #ddd;padding:6px 10px} th{background:#f5f5f5}'
$html = "<!doctype html><html><head><meta charset=\"utf-8\" /><title>公司项目台账报表预览</title><style>$style</style></head><body><h1>公司项目台账报表预览</h1>" + ($sections -join '') + "</body></html>"

$outPath = Join-Path $dir 'index.html'
[System.IO.File]::WriteAllText($outPath, $html, [System.Text.Encoding]::UTF8)
Write-Output ("Generated: " + $outPath)