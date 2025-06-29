# Define task name and description
$taskName = "Run Intune Sync Compliance"

# Intune Detection Script
# This section will exit with 0 if the task exists, which indicates success in Intune
$taskExists = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
if ($null -ne $taskExists) {
    Write-Host "Intune Detection: Task is present. Exiting with code 0."
    exit 0  # Task is present
} else {
    Write-Host "Intune Detection: Task is not present. Exiting with code 1."
    exit 1  # Task is not present
}
