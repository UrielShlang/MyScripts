New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS -ErrorAction SilentlyContinue
$userName = (Get-WmiObject -Class Win32_ComputerSystem | Select-Object -ExpandProperty UserName)
$SID = (New-Object System.Security.Principal.NTAccount($userName)).Translate([System.Security.Principal.SecurityIdentifier]).Value
 reg.exe add "HKU\$SID\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer" /v "NoDesktop" /t REG_DWORD /d 0 /f | Out-Host