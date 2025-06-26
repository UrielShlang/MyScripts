# Outlook for Windows Detection Script for Intune
# Purpose: Detect if Microsoft Outlook for Windows (UWP) is installed
# Exit Code 0: Not found (compliant)
# Exit Code 1: Found (non-compliant)

try {
    # Check for Outlook for Windows packages
    $outlookPackages = Get-AppxPackage -Name "*Microsoft.OutlookForWindows*" -AllUsers -ErrorAction SilentlyContinue
    $provisionedPackages = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Microsoft.OutlookForWindows*" }
    
    $found = $false
    
    if ($outlookPackages) {
        Write-Output "DETECTED: Outlook for Windows installed packages found"
        foreach ($package in $outlookPackages) {
            Write-Output "Package: $($package.PackageFullName)"
        }
        $found = $true
    }
    
    if ($provisionedPackages) {
        Write-Output "DETECTED: Outlook for Windows provisioned packages found"
        foreach ($package in $provisionedPackages) {
            Write-Output "Provisioned: $($package.PackageName)"
        }
        $found = $true
    }
    
    if ($found) {
        Write-Output "STATUS: Outlook for Windows detected - NON-COMPLIANT"
        exit 1
    } else {
        Write-Output "STATUS: Outlook for Windows not found - COMPLIANT"
        exit 0
    }
}
catch {
    Write-Output "ERROR: Detection failed - $($_.Exception.Message)"
    exit 1
}