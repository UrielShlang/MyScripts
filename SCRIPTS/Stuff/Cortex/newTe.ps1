# Collect Antivirus protection data from WMI
$result = @(Get-CimInstance -Namespace 'ROOT\SecurityCenter2' -ClassName AntiVirusProduct)

# Fallback function to find when Cortex last changed/updated local components
Function Get-CortexTimestamp {
    $CortexPaths = @(
        "C:\ProgramData\Cyvera\LocalSystem\Persistence",
        "C:\Program Files\Palo Alto Networks\Cortex XDR"
    )
    ForEach ($Path in $CortexPaths) {
        If (Test-Path $Path) {
            $LatestFile = Get-ChildItem -Path $Path -Recurse -ErrorAction SilentlyContinue | 
                          Sort-Object LastWriteTime -Descending | 
                          Select-Object -First 1
            If ($LatestFile) { return $LatestFile.LastWriteTime }
        }
    }
    return $null
}

# Process the WMI object
$TargetAV = $null
If ($result.count -eq 1) {
    $TargetAV = $result
} ElseIf ($result.count -gt 1) {
    # Prefer the active antivirus if multiple are present on the device
    ForEach ($item in $result) {
        $StateConvert = [System.Convert]::ToString($item.productState,16).padleft(8,'0')
        If ($StateConvert.substring(4,1) -eq '1') {
            $TargetAV = $item
            Break
        }
    }
    if (!$TargetAV) { $TargetAV = $result[-1] }
}

# Build the compliance payload
If ($null -eq $TargetAV) {
    $Output = [PSCustomObject]@{
        AntiVirusProductName = 'No product detected'
        Active               = 'Unknown'
        UptoDate             = $false 
        LastUpdateTime       = 'Unknown'
        IsRecent             = $false
    }
} Else {
    # Ultimate fail-safe: If the product is Cortex, explicitly hardcode the expected compliance string
    if ($TargetAV.displayname -like "*Cortex*") {
        $CleanName = "Cortex XDR Advanced Endpoint Protection"
    } else {
        $CleanName = $TargetAV.displayname -replace '[™®]', ''
    }
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
