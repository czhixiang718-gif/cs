Stop='Stop'
C:\Users\Administrator\Documents\trae_projects\cs\backups = Split-Path -Parent 
 = 'TempRoot!123'
C:\Program Files\MariaDB 12.0\bin\mysqldump.exe = 'C:\Program Files\MariaDB 12.0\bin\mysqldump.exe'
 = Get-Date -Format 'yyyyMMdd_HHmmss'
 = Join-Path C:\Users\Administrator\Documents\trae_projects\cs\backups ("ledger_db_" +  + ".sql")
cmd.exe /c '"' + C:\Program Files\MariaDB 12.0\bin\mysqldump.exe + '" -u root ledger_db > "' +  + '"'
Remove-Item Env:MYSQL_PWD
