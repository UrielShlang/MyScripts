$ExpectedVersion = '32.0.101.7026'
$DeviceNameContains = ''
$ErrorActionPreference = 'SilentlyContinue'

# Retrieve signed drivers for DISPLAY devices
$drivers = Get-CimInstance -ClassName Win32_PnPSignedDriver `
            -Filter "DeviceClass='DISPLAY'" |
            Where-Object {
                ($_.Manufacturer -eq 'Intel Corporation' -or $_.DriverProviderName -eq 'Intel Corporation') `
                -and ($DeviceNameContains -eq '' -or $_.DeviceName -like "*$DeviceNameContains*")
            }

if (-not $drivers) {
    Write-Output "No Intel DISPLAY drivers found."
    exit 1
}

# Get current version (if multiple, take the first unique)
$currentVersion = ($drivers | Select-Object -ExpandProperty DriverVersion -Unique)[0]
$names = ($drivers | Select-Object -ExpandProperty DeviceName -Unique) -join ', '
$dates = ($drivers | Select-Object -ExpandProperty DriverDate -Unique) -join ', '

# Convert versions to System.Version objects for comparison
try {
    $current = [System.Version]$currentVersion
    $expected = [System.Version]$ExpectedVersion
} catch {
    Write-Output "Error comparing versions. Current: $currentVersion; Expected: $ExpectedVersion"
    exit 1
}

# Compare versions
if ($current -lt $expected) {
    # Current version is LOWER than expected – need to install
    Write-Output "Intel display driver version is lower than required. Current: $currentVersion; Expected: $ExpectedVersion. Installation needed on: $names"
    exit 1  # Signal that installation is needed
} else {
    # Current version is EQUAL or HIGHER – no installation needed
    Write-Output "Intel display driver is up to date or newer. Current: $currentVersion; Expected: $ExpectedVersion on: $names (DriverDate: $dates)."
    exit 0  # No installation needed
}
