# If we are running as a 32-bit process on an x64 system, re-launch as a 64-bit process
if ("$env:PROCESSOR_ARCHITEW6432" -ne "ARM64") {
    if (Test-Path "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe") {
      & "$($env:WINDIR)\SysNative\WindowsPowerShell\v1.0\powershell.exe" -ExecutionPolicy bypas
          }
        Exit $lastexitcode
    }

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
