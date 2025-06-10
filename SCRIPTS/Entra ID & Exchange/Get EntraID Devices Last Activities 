Connect-MgGraph -Scopes "Device.Read.All"
Get-MgDevice -All | Select-Object DisplayName,  ApproximateLastSignInDateTime |Export-Csv -Path c:\temp\EntraLastActivity.csv -Encoding utf8
Get-MgDevice -All | Where-Object { $_.OperatingSystem -like "Windows*" } | Select-Object DisplayName, OperatingSystem, ApproximateLastSignInDateTime |Export-Csv -Path c:\temp\EntraLastActivity.csv -Encoding utf8
