#Sample_Secure_Boot_Inventory_Data_Collection_script

# 1. HostName
# PS Version: All | Admin: No | System Requirements: None
try {
    $hostname = $env:COMPUTERNAME
    if ([string]::IsNullOrEmpty($hostname)) {
        Write-Warning "Hostname could not be determined"
        $hostname = "Unknown"
    }
    Write-Host "Hostname: $hostname"
} catch {
    Write-Warning "Error retrieving hostname: $_"
    $hostname = "Error"
    Write-Host "Hostname: $hostname"
}

# 2. CollectionTime
# PS Version: All | Admin: No | System Requirements: None
try {
    $collectionTime = Get-Date
    if ($null -eq $collectionTime) {
        Write-Warning "Could not retrieve current date/time"
        $collectionTime = "Unknown"
    }
    Write-Host "Collection Time: $collectionTime"
} catch {
    Write-Warning "Error retrieving date/time: $_"
    $collectionTime = "Error"
    Write-Host "Collection Time: $collectionTime"
}

# Registry: Secure Boot Main Key (3 values)

# 3. SecureBootEnabled
# PS Version: 3.0+ | Admin: May be required | System Requirements: UEFI/Secure Boot capable system
try {
    $secureBootEnabled = Confirm-SecureBootUEFI -ErrorAction Stop
    Write-Host "Secure Boot Enabled: $secureBootEnabled"
} catch {
    Write-Warning "Unable to determine Secure Boot status via cmdlet: $_"
    # Try registry fallback
    try {
        $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\State" -Name UEFISecureBootEnabled -ErrorAction Stop
        $secureBootEnabled = [bool]$regValue.UEFISecureBootEnabled
        Write-Host "Secure Boot Enabled: $secureBootEnabled"
    } catch {
        Write-Warning "Unable to determine Secure Boot status via registry. System may not support UEFI/Secure Boot."
        $secureBootEnabled = $null
        Write-Host "Secure Boot Enabled: Not Available"
    }
}

# 4. HighConfidenceOptOut
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot" -Name HighConfidenceOptOut -ErrorAction Stop
    $highConfidenceOptOut = $regValue.HighConfidenceOptOut
    Write-Host "High Confidence Opt Out: $highConfidenceOptOut"
} catch {
    Write-Warning "HighConfidenceOptOut registry key not found or inaccessible"
    $highConfidenceOptOut = $null
    Write-Host "High Confidence Opt Out: Not Available"
}

# 5. AvailableUpdates
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot" -Name AvailableUpdates -ErrorAction Stop
    $availableUpdates = $regValue.AvailableUpdates
    if ($null -ne $availableUpdates) {
        # Convert to hexadecimal format
        $availableUpdatesHex = "0x{0:X}" -f $availableUpdates
        Write-Host "Available Updates: $availableUpdatesHex"
    } else {
        Write-Host "Available Updates: Not Available"
    }
} catch {
    Write-Warning "AvailableUpdates registry key not found or inaccessible"
    $availableUpdates = $null
    Write-Host "Available Updates: Not Available"
}

# Registry: Servicing Key (3 values)

# 6. WindowsUEFICA2023Status
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing" -Name UEFICA2023Status -ErrorAction Stop
    $uefica2023Status = $regValue.UEFICA2023Status
    Write-Host "Windows UEFI CA 2023 Status:  $uefica2023Status"
} catch {
    Write-Warning "Windows UEFI CA 2023 Status registry key not found or inaccessible"
    $uefica2023Status = $null
    Write-Host "Windows UEFI CA 2023 Status: Not Available"
}

# 7. WindowsUEFICA2023Capable
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing" -Name WindowsUEFICA2023Capable -ErrorAction Stop
    $WindowsUEFICA2023Capable = $regValue.WindowsUEFICA2023Capable
    Write-Host "Windows UEFI CA 2023 Capable: $Windowsuefica2023Capable"
} catch {
    Write-Warning "WindowsUEFICA2023Capable registry key not found or inaccessible"
    $WindowsUEFICA2023Capable = $null
    Write-Host "Windows UEFI CA 2023 Capable: Not Available"
}

# 8. UEFICA2023Error
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing" -Name UEFICA2023Error -ErrorAction Stop
    $uefica2023Error = $regValue.UEFICA2023Error
    Write-Host "UEFI CA 2023 Error: $uefica2023Error"
} catch {
    Write-Warning "UEFICA2023Error registry key not found or inaccessible"
    $uefica2023Error = $null
    Write-Host "UEFI CA 2023 Error: Not Available"
}

# Registry: Device Attributes (7 values)

# 9. OEMManufacturerName
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes" -Name OEMManufacturerName -ErrorAction Stop
    $oemManufacturerName = $regValue.OEMManufacturerName
    if ([string]::IsNullOrEmpty($oemManufacturerName)) {
        Write-Warning "OEMManufacturerName is empty"
        $oemManufacturerName = "Unknown"
    }
    Write-Host "OEM Manufacturer Name: $oemManufacturerName"
} catch {
    Write-Warning "OEMManufacturerName registry key not found or inaccessible"
    $oemManufacturerName = $null
    Write-Host "OEM Manufacturer Name: Not Available"
}

# 10. OEMModelSystemFamily
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes" -Name OEMModelSystemFamily -ErrorAction Stop
    $oemModelSystemFamily = $regValue.OEMModelSystemFamily
    if ([string]::IsNullOrEmpty($oemModelSystemFamily)) {
        Write-Warning "OEMModelSystemFamily is empty"
        $oemModelSystemFamily = "Unknown"
    }
    Write-Host "OEM Model System Family: $oemModelSystemFamily"
} catch {
    Write-Warning "OEMModelSystemFamily registry key not found or inaccessible"
    $oemModelSystemFamily = $null
    Write-Host "OEM Model System Family: Not Available"
}

# 11. OEMModelNumber
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes" -Name OEMModelNumber -ErrorAction Stop
    $oemModelNumber = $regValue.OEMModelNumber
    if ([string]::IsNullOrEmpty($oemModelNumber)) {
        Write-Warning "OEMModelNumber is empty"
        $oemModelNumber = "Unknown"
    }
    Write-Host "OEM Model Number: $oemModelNumber"
} catch {
    Write-Warning "OEMModelNumber registry key not found or inaccessible"
    $oemModelNumber = $null
    Write-Host "OEM Model Number: Not Available"
}

# 12. FirmwareVersion
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes" -Name FirmwareVersion -ErrorAction Stop
    $firmwareVersion = $regValue.FirmwareVersion
    if ([string]::IsNullOrEmpty($firmwareVersion)) {
        Write-Warning "FirmwareVersion is empty"
        $firmwareVersion = "Unknown"
    }
    Write-Host "Firmware Version: $firmwareVersion"
} catch {
    Write-Warning "FirmwareVersion registry key not found or inaccessible"
    $firmwareVersion = $null
    Write-Host "Firmware Version: Not Available"
}

# 13. FirmwareReleaseDate
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes" -Name FirmwareReleaseDate -ErrorAction Stop
    $firmwareReleaseDate = $regValue.FirmwareReleaseDate
    if ([string]::IsNullOrEmpty($firmwareReleaseDate)) {
        Write-Warning "FirmwareReleaseDate is empty"
        $firmwareReleaseDate = "Unknown"
    }
    Write-Host "Firmware Release Date: $firmwareReleaseDate"
} catch {
    Write-Warning "FirmwareReleaseDate registry key not found or inaccessible"
    $firmwareReleaseDate = $null
    Write-Host "Firmware Release Date: Not Available"
}

# 14. OSArchitecture
# PS Version: All | Admin: No | System Requirements: None
try {
    $osArchitecture = $env:PROCESSOR_ARCHITECTURE
    if ([string]::IsNullOrEmpty($osArchitecture)) {
        # Try registry fallback
        $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes" -Name OSArchitecture -ErrorAction Stop
        $osArchitecture = $regValue.OSArchitecture
    }
    if ([string]::IsNullOrEmpty($osArchitecture)) {
        Write-Warning "OSArchitecture could not be determined"
        $osArchitecture = "Unknown"
    }
    Write-Host "OS Architecture: $osArchitecture"
} catch {
    Write-Warning "Error retrieving OSArchitecture: $_"
    $osArchitecture = "Unknown"
    Write-Host "OS Architecture: $osArchitecture"
}

# 15. CanAttemptUpdateAfter (FILETIME)
# PS Version: All | Admin: May be required | System Requirements: None
try {
    $regValue = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\SecureBoot\Servicing\DeviceAttributes" -Name CanAttemptUpdateAfter -ErrorAction Stop
    $canAttemptUpdateAfter = $regValue.CanAttemptUpdateAfter
    # Convert FILETIME to DateTime if it's a valid number
    if ($null -ne $canAttemptUpdateAfter -and $canAttemptUpdateAfter -is [long]) {
        try {
            $canAttemptUpdateAfter = [DateTime]::FromFileTime($canAttemptUpdateAfter)
        } catch {
            Write-Warning "Could not convert CanAttemptUpdateAfter FILETIME to DateTime"
        }
    }
    Write-Host "Can Attempt Update After: $canAttemptUpdateAfter"
} catch {
    Write-Warning "CanAttemptUpdateAfter registry key not found or inaccessible"
    $canAttemptUpdateAfter = $null
    Write-Host "Can Attempt Update After: Not Available"
}

# Event Logs: System Log (5 values)

# 16-20. Event Log queries
# PS Version: 3.0+ | Admin: May be required for System log | System Requirements: None
try {
    $allEventIds = @(1801, 1808)
    $events = @(Get-WinEvent -FilterHashtable @{LogName='System'; ID=$allEventIds} -MaxEvents 20 -ErrorAction Stop)

    if ($events.Count -eq 0) {
        Write-Warning "No Secure Boot events (1801/1808) found in System log"
        $latestEventId = $null
        $bucketId = $null
        $confidence = $null
        $event1801Count = 0
        $event1808Count = 0
        Write-Host "Latest Event ID: Not Available"
        Write-Host "Bucket ID: Not Available"
        Write-Host "Confidence: Not Available"
        Write-Host "Event 1801 Count: 0"
        Write-Host "Event 1808 Count: 0"
    } else {
        # 16. LatestEventId
        $latestEvent = $events | Sort-Object TimeCreated -Descending | Select-Object -First 1
        if ($null -eq $latestEvent) {
            Write-Warning "Could not determine latest event"
            $latestEventId = $null
            Write-Host "Latest Event ID: Not Available"
        } else {
            $latestEventId = $latestEvent.Id
            Write-Host "Latest Event ID: $latestEventId"
        }

        # 17. BucketID - Extracted from Event 1801/1808
        if ($null -ne $latestEvent -and $null -ne $latestEvent.Message) {
            if ($latestEvent.Message -match 'BucketId:\s*(.+)') {
                $bucketId = $matches[1].Trim()
                Write-Host "Bucket ID: $bucketId"
            } else {
                Write-Warning "BucketId not found in event message"
                $bucketId = $null
                Write-Host "Bucket ID: Not Found in Event"
            }
        } else {
            Write-Warning "Latest event or message is null, cannot extract BucketId"
            $bucketId = $null
            Write-Host "Bucket ID: Not Available"
        }

        # 18. Confidence - Extracted from Event 1801/1808
        if ($null -ne $latestEvent -and $null -ne $latestEvent.Message) {
            if ($latestEvent.Message -match 'BucketConfidenceLevel:\s*(.+)') {
                $confidence = $matches[1].Trim()
                Write-Host "Confidence: $confidence"
            } else {
                Write-Warning "Confidence level not found in event message"
                $confidence = $null
                Write-Host "Confidence: Not Found in Event"
            }
        } else {
            Write-Warning "Latest event or message is null, cannot extract Confidence"
            $confidence = $null
            Write-Host "Confidence: Not Available"
        }

        # 19. Event1801Count
        $event1801Array = @($events | Where-Object {$_.Id -eq 1801})
        $event1801Count = $event1801Array.Count
        Write-Host "Event 1801 Count: $event1801Count"

        # 20. Event1808Count
        $event1808Array = @($events | Where-Object {$_.Id -eq 1808})
        $event1808Count = $event1808Array.Count
        Write-Host "Event 1808 Count: $event1808Count"
    }
} catch {
    Write-Warning "Error retrieving event logs. May require administrator privileges: $_"
    $latestEventId = $null
    $bucketId = $null
    $confidence = $null
    $event1801Count = 0
    $event1808Count = 0
    Write-Host "Latest Event ID: Error"
    Write-Host "Bucket ID: Error"
    Write-Host "Confidence: Error"
    Write-Host "Event 1801 Count: 0"
    Write-Host "Event 1808 Count: 0"
}

# WMI/CIM Queries (4 values)

# 21. OSVersion
# PS Version: 3.0+ (use Get-WmiObject for 2.0) | Admin: No | System Requirements: None
try {
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    if ($null -eq $osInfo -or [string]::IsNullOrEmpty($osInfo.Version)) {
        Write-Warning "Could not retrieve OS version"
        $osVersion = "Unknown"
    } else {
        $osVersion = $osInfo.Version
    }
    Write-Host "OS Version: $osVersion"
} catch {
    Write-Warning "Error retrieving OS version: $_"
    $osVersion = "Unknown"
    Write-Host "OS Version: $osVersion"
}

# 22. LastBootTime
# PS Version: 3.0+ (use Get-WmiObject for 2.0) | Admin: No | System Requirements: None
try {
    $osInfo = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
    if ($null -eq $osInfo -or $null -eq $osInfo.LastBootUpTime) {
        Write-Warning "Could not retrieve last boot time"
        $lastBootTime = $null
        Write-Host "Last Boot Time: Not Available"
    } else {
        $lastBootTime = $osInfo.LastBootUpTime
        Write-Host "Last Boot Time: $lastBootTime"
    }
} catch {
    Write-Warning "Error retrieving last boot time: $_"
    $lastBootTime = $null
    Write-Host "Last Boot Time: Not Available"
}

# 23. BaseBoardManufacturer
# PS Version: 3.0+ (use Get-WmiObject for 2.0) | Admin: No | System Requirements: None
try {
    $baseBoard = Get-CimInstance Win32_BaseBoard -ErrorAction Stop
    if ($null -eq $baseBoard -or [string]::IsNullOrEmpty($baseBoard.Manufacturer)) {
        Write-Warning "Could not retrieve baseboard manufacturer"
        $baseBoardManufacturer = "Unknown"
    } else {
        $baseBoardManufacturer = $baseBoard.Manufacturer
    }
    Write-Host "Baseboard Manufacturer: $baseBoardManufacturer"
} catch {
    Write-Warning "Error retrieving baseboard manufacturer: $_"
    $baseBoardManufacturer = "Unknown"
    Write-Host "Baseboard Manufacturer: $baseBoardManufacturer"
}

# 24. BaseBoardProduct
# PS Version: 3.0+ (use Get-WmiObject for 2.0) | Admin: No | System Requirements: None
try {
    $baseBoard = Get-CimInstance Win32_BaseBoard -ErrorAction Stop
    if ($null -eq $baseBoard -or [string]::IsNullOrEmpty($baseBoard.Product)) {
        Write-Warning "Could not retrieve baseboard product"
        $baseBoardProduct = "Unknown"
    } else {
        $baseBoardProduct = $baseBoard.Product
    }
    Write-Host "Baseboard Product: $baseBoardProduct"
} catch {
    Write-Warning "Error retrieving baseboard product: $_"
    $baseBoardProduct = "Unknown"
    Write-Host "Baseboard Product: $baseBoardProduct"
    }