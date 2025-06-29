# Define the task action
$action = New-ScheduledTaskAction -Execute "C:\Program Files (x86)\Microsoft Intune Management Extension\Microsoft.Management.Services.IntuneWindowsAgent.exe" -Argument "intunemanagementextension://synccompliance"

# Define the task trigger to run every 5 minutes
$startBoundary = (Get-Date).AddMinutes(1) # Start 1 minute from now to avoid "past start time" issues
$trigger = New-ScheduledTaskTrigger -Once -At $startBoundary -RepetitionInterval (New-TimeSpan -Minutes 5) -RepetitionDuration (New-TimeSpan -Days (365 * 20))

# Define the task settings
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -DontStopOnIdleEnd -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

# Define the principal (running as SYSTEM)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount

# Register the scheduled task
Register-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal -TaskName "Run Intune Sync Compliance" -Description "Runs the Intune Sync Compliance every 5 minutes"

Write-Host "Scheduled Task Created Successfully"
