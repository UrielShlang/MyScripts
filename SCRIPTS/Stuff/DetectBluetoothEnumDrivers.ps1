$ErrorActionPreference = 'SilentlyContinue'

# Define the required Bluetooth Enumerator drivers
$RequiredDrivers = @(
    'Microsoft Bluetooth LE Enumerator',
    'Microsoft Bluetooth Enumerator'
)

# Retrieve all Bluetooth drivers
$drivers = Get-CimInstance -ClassName Win32_PnPSignedDriver -Filter "DeviceClass='BLUETOOTH'"

if (-not $drivers) {
    Write-Output "No Bluetooth drivers found on the system."
    exit 1
}

# Check for each required driver
$foundDrivers = @()
$missingDrivers = @()

foreach ($requiredDriver in $RequiredDrivers) {
    $match = $drivers | Where-Object { $_.DeviceName -eq $requiredDriver }
    
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
