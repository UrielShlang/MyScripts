$ExpectedVersion = '23.40.0.2'
# Optional: Filter by device name (e.g., "Realtek Bluetooth"). If empty – all Realtek Bluetooth drivers will be checked.
$DeviceNameContains = ''
$ErrorActionPreference = 'SilentlyContinue'

# Retrieve signed drivers for Bluetooth devices
$drivers = Get-CimInstance -ClassName Win32_PnPSignedDriver `
           -Filter "DeviceClass='Bluetooth'" |
           Where-Object {
               ($_.Manufacturer -like '*Intel*' -or $_.DriverProviderName -like '*Intel*') `
               -and ($DeviceNameContains -eq '' -or $_.DeviceName -like "*$DeviceNameContains*")
           }

if (-not $drivers) {
    Write-Output "No Intel Bluetooth drivers found."
    exit 1
}

# Check for exact version match
$match = $drivers | Where-Object { $_.DriverVersion -eq $ExpectedVersion }

if ($match) {
    # Informational output, but most importantly – Exit 0 for Intune detection
    $names = ($match | Select-Object -ExpandProperty DeviceName -Unique) -join ', '
    $vers  = ($match | Select-Object -ExpandProperty DriverVersion -Unique) -join ', '
    $dates = ($match | Select-Object -ExpandProperty DriverDate -Unique) -join ', '
    Write-Output "Detected Intel Bluetooth driver $vers on: $names (DriverDate: $dates)."
    exit 0
}
else {
    $current = ($drivers | Select-Object -ExpandProperty DriverVersion -Unique) -join ', '
    Write-Output "Intel Bluetooth driver found, but version is different. Current: $current; Expected: $ExpectedVersion"
    exit 1
}
