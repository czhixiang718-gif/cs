Stop='Stop'
C:\Users\Administrator\Documents\trae_projects\cs\exports = Split-Path -Parent 
 = '8f7ee940d85446218a65ac7941fa4014f53ab2a0e7cf4484806944117b8cefb7'
C:\Program Files\MariaDB 12.0\bin\mysql.exe = 'C:\Program Files\MariaDB 12.0\bin\mysql.exe'
ledger_db = 'ledger_db'
function Export-MySqlCsv([string],[string]){
   = & C:\Program Files\MariaDB 12.0\bin\mysql.exe -u ledger_app -D ledger_db -B -e 
   =  | ConvertFrom-Csv -Delimiter "	"
   | Export-Csv -Path  -NoTypeInformation -Encoding UTF8
}
Export-MySqlCsv 'SELECT * FROM monthly_contracts' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'monthly_contracts.csv')
Export-MySqlCsv 'SELECT * FROM monthly_payments' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'monthly_payments.csv')
Export-MySqlCsv 'SELECT * FROM project_completion_rank' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'project_completion_rank.csv')
Export-MySqlCsv 'SELECT * FROM project_finance_summary' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'project_finance_summary.csv')
Export-MySqlCsv 'SELECT * FROM global_finance_totals' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'global_finance_totals.csv')
Export-MySqlCsv 'SELECT * FROM quarterly_contracts' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'quarterly_contracts.csv')
Export-MySqlCsv 'SELECT * FROM quarterly_payments' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'quarterly_payments.csv')
Export-MySqlCsv 'SELECT * FROM yearly_contracts' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'yearly_contracts.csv')
Export-MySqlCsv 'SELECT * FROM yearly_payments' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'yearly_payments.csv')
Export-MySqlCsv 'SELECT * FROM overdue_projects_90d' (Join-Path C:\Users\Administrator\Documents\trae_projects\cs\exports 'overdue_projects_90d.csv')
Remove-Item Env:MYSQL_PWD
