# Task settings
$taskName = "UpdateDefenderSignaturesOnUnlock"
$taskDescription = "Checks and updates Windows Defender signatures when the workstation is unlocked."

# Define the action
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-Command `"if ((Get-MpComputerStatus).DefenderSignaturesOutOfDate) { Update-MpSignature; Write-Output 'Signatures updated.' } else { Write-Output 'Signatures are up-to-date.' }`""

# Define the trigger for "On workstation unlock"
$trigger = New-ScheduledTaskTrigger -AtWorkstationUnlock

# Define the principal (run whether user is logged on or not)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Define additional task settings
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -Hidden

# Register the task
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Description $taskDescription

Write-Output "Task '$taskName' has been created successfully."
