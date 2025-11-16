$ErrorActionPreference = 'SilentlyContinue'

# Define the required Bluetooth Enumerator drivers
$RequiredDrivers = @(
    'Microsoft Bluetooth LE Enumerator',
    'Microsoft Bluetooth Enumerator'
)

# Retrieve all PnP devices with Present status (actually installed and available)
$devices = Get-PnpDevice | Where-Object { $_.Status -eq 'OK' -or $_.Status -eq 'Degraded' }

if (-not $devices) {
    Write-Output "No active devices found on the system."
    exit 1
}

# Check for each required driver
$foundDrivers = @()
$missingDrivers = @()

foreach ($requiredDriver in $RequiredDrivers) {
    $match = $devices | Where-Object { $_.FriendlyName -eq $requiredDriver }
    
    if ($match) {
        $foundDrivers += $requiredDriver
        Write-Output "Found: $requiredDriver (Status: $($match.Status))"
    }
    else {
        $missingDrivers += $requiredDriver
    }
}

# Determine exit code based on results
if ($missingDrivers.Count -eq 0) {
    Write-Output "All required Bluetooth Enumerator drivers are installed and active."
    exit 0
}
else {
    Write-Output "Missing or inactive Bluetooth Enumerator drivers: $($missingDrivers -join ', ')"
    exit 1
}
