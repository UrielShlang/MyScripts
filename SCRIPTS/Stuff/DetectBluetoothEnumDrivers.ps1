$ErrorActionPreference = 'SilentlyContinue'

# Define the required Bluetooth Enumerator drivers
$RequiredDrivers = @(
    'Microsoft Bluetooth LE Enumerator',
    'Microsoft Bluetooth Enumerator'
)

# Retrieve all PnP devices (not just signed drivers)
$devices = Get-CimInstance -ClassName Win32_PnPEntity | 
           Where-Object { $_.ConfigManagerErrorCode -ne 22 }  # 22 = Device is disabled

if (-not $devices) {
    Write-Output "No devices found on the system."
    exit 1
}

# Check for each required driver
$foundDrivers = @()
$missingDrivers = @()

foreach ($requiredDriver in $RequiredDrivers) {
    $match = $devices | Where-Object { $_.Name -eq $requiredDriver }
    
    if ($match) {
        $foundDrivers += $requiredDriver
    }
    else {
        $missingDrivers += $requiredDriver
    }
}

# Determine exit code based on results
if ($missingDrivers.Count -eq 0) {
    Write-Output "All required Bluetooth Enumerator drivers are installed: $($foundDrivers -join ', ')"
    exit 0
}
else {
    Write-Output "Missing Bluetooth Enumerator drivers: $($missingDrivers -join ', ')"
    if ($foundDrivers.Count -gt 0) {
        Write-Output "Found drivers: $($foundDrivers -join ', ')"
    }
    exit 1
}
