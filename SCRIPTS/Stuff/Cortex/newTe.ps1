
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
