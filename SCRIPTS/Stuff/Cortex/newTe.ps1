# Parse updates and timestamps

$TargetDate = if ($TargetAV.timestamp -is [System.DateTime]) { $TargetAV.timestamp } else { Get-CortexTimestamp }

If ($TargetDate) {

$LastUpdateTime = Get-Date $TargetDate -Format "yyyy-MM-dd HH:mm:ss"

$IsRecentString = if ($TargetDate -gt (Get-Date).AddDays(-7)) { 'True' } else { 'False' }

} Else {

$LastUpdateTime = 'Unknown'

$IsRecentString = 'False'

}

# Parse Product State Flag

$StateConvert = [System.Convert]::ToString($TargetAV.productState,16).padleft(8,'0')

$Active = Switch ($StateConvert.substring(4,1)) {

'0' {'Off'}

'1' {'On'}

'2' {'Snoozed'}

'3' {'Expired'}

Default {'Unknown'}

}

# FIX 2: Map UptoDate strictly to a Boolean datatype ($true / $false)

$UptoDate = Switch ($StateConvert.substring(6,1)) {

'0' { $true }

'1' { $false }

Default { $false }

}

# FIX 3: Map IsRecent strictly to a Boolean datatype ($true / $false)

$IsRecent = if ($IsRecentString -eq 'True') { $true } else { $false }

# Construct final object matching Intune property casings and expected types

$Output = [PSCustomObject]@{

AntiVirusProductName = $CleanName

Active = $Active

UptoDate = $UptoDate

LastUpdateTime = $LastUpdateTime

IsRecent = $IsRecent

ScriptVersion = "2.0.1-NoXDRT" # <--- Update this when testing changes > logs show results "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\AgentExecutor.log"

}

}

# Compress into clean JSON payload for the Intune Management Extension agent

$ThresholdOutput = $Output | ConvertTo-Json -Compress

Write-Output $ThresholdOutput
