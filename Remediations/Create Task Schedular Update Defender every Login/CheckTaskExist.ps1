###
###
###

$TaskName = "UpdateDefenderSignaturesOnUnlock"
$TaskExists = Get-ScheduledTask | Where-Object { $_.TaskName -eq $TaskName }

if ($TaskExists) {
    Write-Output "The scheduled task '$TaskName' exists."
    exit 0
} else {
    Write-Output "The scheduled task '$TaskName' does not exist."
    exit 1
}
