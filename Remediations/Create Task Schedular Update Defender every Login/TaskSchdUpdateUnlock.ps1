
####
####
#### Check if the script is running in a 32-bit process on a 64-bit OS
$is64BitOS = [Environment]::Is64BitOperatingSystem
$is64BitProcess = [Environment]::Is64BitProcess

if ($is64BitOS -and -not $is64BitProcess) {
    # Relaunch the script in a 64-bit PowerShell process
    Write-Host "Relaunching script in 64-bit PowerShell process..."
    $scriptPath = $MyInvocation.MyCommand.Path
    $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`""
    Write-Host "Correct with change Environment to 64 bit" -ForegroundColor Green
    Start-Process -FilePath "$env:windir\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList $arguments -Wait
    
    Exit $lastexitcode
}
$Shell = New-Object -ComObject "WScript.Shell"
$Button = $Shell.Popup("Click OK to continue.", 0, "Hello", 0)

# Task settings
$taskName = "UpdateDefenderSignaturesOnUnlock"
$taskDescription = "Checks and updates Windows Defender signatures when the workstation is unlocked."

# Define the action
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-Command `"Update-MpSignature`""

# Define the trigger for "On workstation unlock"

$stateChangeTrigger = Get-CimClass `
    -Namespace ROOT\Microsoft\Windows\TaskScheduler `
    -ClassName MSFT_TaskSessionStateChangeTrigger

$onUnlockTrigger = New-CimInstance `
    -CimClass $stateChangeTrigger `
    -Property @{
        StateChange = 8  # TASK_SESSION_STATE_CHANGE_TYPE.TASK_SESSION_UNLOCK (taskschd.h)
    } `
    -ClientOnly

#$trigger = New-ScheduledTaskTrigger $onUnlockTrigger

# Define the principal (run whether user is logged on or not)
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Define additional task settings
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -Hidden

# Register the task
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $onUnlockTrigger -Principal $principal -Settings $settings -Description $taskDescription

Write-Output "Task '$taskName' has been created successfully."

