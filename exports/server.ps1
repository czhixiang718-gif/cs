$ErrorActionPreference='Stop'
$port = 8000
$prefix = "http://localhost:$port/"
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add($prefix)
$listener.Start()
Write-Host "Preview: $prefix"
$exports = Join-Path (Split-Path -Parent $PSCommandPath) '.'
function HtmlTable([string]$path,[string]$title){
  if (Test-Path $path) {
    $rows = Import-Csv -Path $path
    $cols = @()
    if ($rows.Count -gt 0) { $cols = $rows[0].PSObject.Properties.Name }
    $sb = "<h2>$title</h2><table border='1'><tr>"
    foreach ($c in $cols) { $sb += "<th>$c</th>" }
    $sb += "</tr>"
    foreach ($r in $rows) {
      $sb += "<tr>"
      foreach ($c in $cols) { $sb += "<td>" + ($r.$c) + "</td>" }
      $sb += "</tr>"
    }
    $sb += "</table>"
    return $sb
  } else {
    return "<h2>$title</h2><p>文件未找到</p>"
  }
}
while ($true) {
  $ctx = $listener.GetContext()
  $res = $ctx.Response
  try {
    $dir = "$exports"
    $html = "<html><head><meta charset='utf-8'><title>台账预览</title><style>body{font-family:Segoe UI,Arial;margin:20px} table{border-collapse:collapse} th,td{padding:6px 10px}</style></head><body><h1>公司项目台账报表预览</h1>"
    $html += HtmlTable (Join-Path $dir 'global_finance_totals.csv') '全量合计'
    $html += HtmlTable (Join-Path $dir 'project_finance_summary.csv') '项目汇总明细'
    $html += HtmlTable (Join-Path $dir 'monthly_contracts.csv') '月签约'
    $html += HtmlTable (Join-Path $dir 'monthly_payments.csv') '月回款'
    $html += HtmlTable (Join-Path $dir 'project_completion_rank.csv') '完成率排行'
    $html += HtmlTable (Join-Path $dir 'quarterly_contracts.csv') '季度签约'
    $html += HtmlTable (Join-Path $dir 'quarterly_payments.csv') '季度回款'
    $html += HtmlTable (Join-Path $dir 'yearly_contracts.csv') '年度签约'
    $html += HtmlTable (Join-Path $dir 'yearly_payments.csv') '年度回款'
    $html += HtmlTable (Join-Path $dir 'overdue_projects_90d.csv') '逾期未回款&gt;90天'
    $html += "</body></html>"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($html)
    $res.ContentType = "text/html; charset=utf-8"
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
  } finally {
    $res.Close()
  }
}
