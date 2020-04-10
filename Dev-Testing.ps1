
 Import-ExcelSched -Path '.\CWO\CWO 20-06 Schedule.xlsx'  -Course "CVAH" -Class "20-04" | Export-Csv .\events.csv -Force
 Import-Csv .\events.csv | Out-GridView