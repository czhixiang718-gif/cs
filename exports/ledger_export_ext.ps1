$ErrorActionPreference='Stop'
$dir = Split-Path -Parent $PSCommandPath
$env:MYSQL_PWD = 'Report!123'
$bin = 'C:\Program Files\MariaDB 12.0\bin\mysql.exe'
$db = 'ledger_db'
$pw = 'Report!123'
function Export-MySqlCsv([string]$query,[string]$outfile){
  $tsv = & $bin -u ledger_report -D $db --password=$pw -B -e $query
  $csvObjs = $tsv | ConvertFrom-Csv -Delimiter "`t"
  $csvObjs | Export-Csv -Path $outfile -NoTypeInformation -Encoding UTF8
}
Export-MySqlCsv 'SELECT * FROM monthly_contracts' (Join-Path $dir 'monthly_contracts.csv')
Export-MySqlCsv 'SELECT * FROM monthly_payments' (Join-Path $dir 'monthly_payments.csv')
Export-MySqlCsv 'SELECT * FROM project_completion_rank' (Join-Path $dir 'project_completion_rank.csv')
Export-MySqlCsv 'SELECT * FROM project_finance_summary' (Join-Path $dir 'project_finance_summary.csv')
Export-MySqlCsv 'SELECT * FROM global_finance_totals' (Join-Path $dir 'global_finance_totals.csv')
Export-MySqlCsv 'SELECT * FROM quarterly_contracts' (Join-Path $dir 'quarterly_contracts.csv')
Export-MySqlCsv 'SELECT * FROM quarterly_payments' (Join-Path $dir 'quarterly_payments.csv')
Export-MySqlCsv 'SELECT * FROM yearly_contracts' (Join-Path $dir 'yearly_contracts.csv')
Export-MySqlCsv 'SELECT * FROM yearly_payments' (Join-Path $dir 'yearly_payments.csv')
Export-MySqlCsv 'SELECT * FROM overdue_projects_90d' (Join-Path $dir 'overdue_projects_90d.csv')
Remove-Item Env:MYSQL_PWD