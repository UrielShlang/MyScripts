$AVClient = 'Cortex XDR'
$ServiceName = 'cyserver'
$ExpectedPath = 'C:\Program Files\Palo Alto Networks\Traps\cyserver.exe'
$AVSummary = @{}

# Check WSC AV entry (Cortex XDR may register as "Cortex XDR", "Palo Alto Networks Cortex XDR" or "Traps")
$AVProduct = Get-CimInstance -Namespace 'root\SecurityCenter2' -ClassName AntiVirusProduct -ErrorAction SilentlyContinue |
    Where-Object { $_.displayName -match 'Cortex\s*XDR|Palo\s*Alto|Traps' } |
    Select-Object -First 1

# Check service
$cxService = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue

# File check
$cxBinaryExists = Test-Path $ExpectedPath

# Initialize fields
$AVSummary["$AVClient WSC detected"] = if ($AVProduct) { "Yes" } else { "No" }
$AVSummary["$AVClient service running"] = if ($cxService -and $cxService.Status -eq 'Running') { "Yes" } else { "No" }
$AVSummary["$AVClient binary present"] = if ($cxBinaryExists) { "Yes" } else { "No" }

# Optional: Add WSC-based protection status if available
if ($AVProduct) {
    $hexProductState = [Convert]::ToString($AVProduct.productState, 16).PadLeft(6, '0')
    $hexRealTimeProtection = $hexProductState.Substring(2, 2)
    $hexDefinitionStatus = $hexProductState.Substring(4, 2)

    $RealTimeProtectionStatus = switch ($hexRealTimeProtection) {
        '00' { 'Off' }
        '01' { 'Expired' }
        '10' { 'On' }
        '11' { 'Snoozed' }
        default { 'Unknown' }
    }

    $DefinitionStatus = switch ($hexDefinitionStatus) {
        '00' { 'Up to Date' }
        '10' { 'Out of Date' }
        default { 'Unknown' }
    }

    $AVSummary["$AVClient real time protection enabled"] = $RealTimeProtectionStatus
    $AVSummary["$AVClient definitions up-to-date"] = $DefinitionStatus
}
else {
    $AVSummary["$AVClient real time protection enabled"] = 'Unknown'
    $AVSummary["$AVClient definitions up-to-date"] = 'Unknown'
}

Return $AVSummary | ConvertTo-Json -Compress
