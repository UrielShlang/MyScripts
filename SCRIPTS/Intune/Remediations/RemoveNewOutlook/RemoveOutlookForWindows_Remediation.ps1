# Outlook for Windows Remediation Script for Intune
# Purpose: Remove Microsoft Outlook for Windows (UWP) packages
# Exit Code 0: Success
# Exit Code 1: Failure

try {
    Write-Output "Starting Outlook for Windows removal process..."
    
    $removalSuccess = $true
    $packagesRemoved = 0
    
    # Method 1: Remove installed packages for all users
    $outlookPackages = Get-AppxPackage -Name "*Microsoft.OutlookForWindows*" -AllUsers -ErrorAction SilentlyContinue
    
    if ($outlookPackages) {
        Write-Output "Found $($outlookPackages.Count) installed Outlook for Windows package(s)"
        
        foreach ($package in $outlookPackages) {
            try {
                Write-Output "Removing package: $($package.PackageFullName)"
                Remove-AppxPackage -Package $package.PackageFullName -AllUsers -ErrorAction Stop
                Write-Output "Successfully removed: $($package.PackageFullName)"
                $packagesRemoved++
            }
            catch {
                Write-Output "Failed to remove package: $($package.PackageFullName) - Error: $($_.Exception.Message)"
                $removalSuccess = $false
            }
        }
    }
    
    # Method 2: Remove provisioned packages
    $provisionedPackages = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Microsoft.OutlookForWindows*" }
    
    if ($provisionedPackages) {
        Write-Output "Found $($provisionedPackages.Count) provisioned Outlook for Windows package(s)"
        
        foreach ($package in $provisionedPackages) {
            try {
                Write-Output "Removing provisioned package: $($package.PackageName)"
                Remove-AppxProvisionedPackage -Online -PackageName $package.PackageName -ErrorAction Stop
                Write-Output "Successfully removed provisioned package: $($package.PackageName)"
                $packagesRemoved++
            }
            catch {
                Write-Output "Failed to remove provisioned package: $($package.PackageName) - Error: $($_.Exception.Message)"
                $removalSuccess = $false
            }
        }
    }
    
    # Method 3: Alternative removal using the original command approach
    try {
        $outlookPackage = Get-AppxPackage Microsoft.OutlookForWindows -ErrorAction SilentlyContinue
        if ($outlookPackage) {
            Write-Output "Attempting alternative removal method..."
            Remove-AppxProvisionedPackage -AllUsers -Online -PackageName $outlookPackage.PackageFullName -ErrorAction Stop
            Write-Output "Alternative removal method succeeded"
            $packagesRemoved++
        }
    }
    catch {
        Write-Output "Alternative removal method failed: $($_.Exception.Message)"
    }
    
    # Verification
    Start-Sleep -Seconds 3
    
    $remainingPackages = Get-AppxPackage -Name "*Microsoft.OutlookForWindows*" -AllUsers -ErrorAction SilentlyContinue
    $remainingProvisioned = Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Microsoft.OutlookForWindows*" }
    
    if ($remainingPackages -or $remainingProvisioned) {
        Write-Output "WARNING: Some Outlook for Windows packages may still remain"
        if ($remainingPackages) {
            foreach ($pkg in $remainingPackages) {
                Write-Output "Remaining package: $($pkg.PackageFullName)"
            }
        }
        if ($remainingProvisioned) {
            foreach ($pkg in $remainingProvisioned) {
                Write-Output "Remaining provisioned: $($pkg.PackageName)"
            }
        }
        $removalSuccess = $false
    }
    
    if ($packagesRemoved -eq 0) {
        Write-Output "No Outlook for Windows packages found to remove"
        exit 0
    }
    
    if ($removalSuccess) {
        Write-Output "SUCCESS: Removed $packagesRemoved Outlook for Windows package(s)"
        Write-Output "Remediation completed successfully"
        exit 0
    } else {
        Write-Output "PARTIAL SUCCESS: Removed $packagesRemoved package(s) but some issues occurred"
        exit 1
    }
}
catch {
    Write-Output "CRITICAL ERROR: Remediation script failed - $($_.Exception.Message)"
    exit 1
}