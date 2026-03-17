# Detect Laptop/Desktop by ChassisType (Win32_SystemEnclosure)
# Run as standard user (admin not required)

$enc = Get-CimInstance -ClassName Win32_SystemEnclosure -ErrorAction Stop
$types = @($enc.ChassisTypes | Where-Object { $_ -ne $null })

$map = @{
  3  = "Desktop"
  4  = "Low Profile Desktop"
  5  = "Pizza Box"
  6  = "Mini Tower"
  7  = "Tower"
  8  = "Portable"
  9  = "Laptop"
  10 = "Notebook"
  11 = "Hand Held"
  12 = "Docking Station"
  13 = "All-in-One"
  14 = "Sub Notebook"
  15 = "Space-Saving"
  16 = "Lunch Box"
  17 = "Main System Chassis"
  18 = "Expansion Chassis"
  21 = "Peripheral Chassis"
  23 = "Rack Mount Chassis"
  30 = "Tablet"
  31 = "Convertible"
  32 = "Detachable"
}

$laptopCodes  = 8,9,10,14,30,31,32
$desktopCodes = 3,4,5,6,7,15,16,17,23

$labels = $types | ForEach-Object {
  [PSCustomObject]@{
    Code  = $_
    Label = ($map[$_] ?? "Unknown")
  }
}

$classification =
  if ($types.Count -eq 0) { "Unknown (no ChassisTypes reported)" }
  elseif ($types | Where-Object { $_ -in $laptopCodes }) { "Laptop" }
  elseif ($types | Where-Object { $_ -in $desktopCodes }) { "Desktop" }
  else { "Unknown/Other" }

$cs = Get-CimInstance -ClassName Win32_ComputerSystem

[PSCustomObject]@{
  ComputerName   = $env:COMPUTERNAME
  Manufacturer   = $cs.Manufacturer
  Model          = $cs.Model
  ChassisTypes   = ($labels | ForEach-Object { "$($_.Code)=$($_.Label)" }) -join ", "
  Classification = $classification
} | Format-List
