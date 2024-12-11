# Script for uninstalling and reinstalling Citrix Workspace App

# Configuration for the installer file
$installerPath = ".\CitrixWorkspaceFullInstaller.exe"  # Update the path to the installer file

# Check if the installer file exists
if (!(Test-Path -Path $installerPath)) {
    Write-Host "The installer file was not found at: $installerPath"
    exit 1
}

# Silent uninstall of Citrix Workspace App
Write-Host "Performing a silent uninstall of Citrix Workspace App..."
Start-Process -FilePath $installerPath -ArgumentList "/uninstall /silent /noreboot" -NoNewWindow -Wait

# Check if the operation completed successfully
if ($?) {
    Write-Host "Uninstall completed successfully."
} else {
    Write-Host "An error occurred during uninstallation."
    exit 1
}

# Reinstall with Force Install and Clean Install
Write-Host "Starting reinstallation of Citrix Workspace App..."
Start-Process -FilePath $installerPath -ArgumentList "/silent /forceinstall /noreboot /CleanInstall /AutoUpdateCheck=disabled /includeSSON /ENABLE_SSON=Yes /ALLOWADDSTORE=A /ALLOWSAVEPWD=N" -NoNewWindow -Wait

# Check if the operation completed successfully
if ($?) {
    Write-Host "Installation completed successfully."
} else {
    Write-Host "An error occurred during installation."
    exit 1
}

Write-Host "The process completed successfully!"
