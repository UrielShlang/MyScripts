# Script for uninstalling and reinstalling Citrix Workspace App

# Configuration for the installer file and download URL
$installerPath = "./CitrixWorkspaceFullInstaller.exe"  # Update the path to the installer file
$downloadUrl = "https://downloads.citrix.com/22865/CitrixWorkspaceFullInstaller.exe?__gda__=exp=1733919746~acl=/*~hmac=64895eb080a3a370ac40100da3da16c4c075ab04cedebfd33686d92d8311be40"

# Function to download the installer if it does not exist
function Download-Installer {
    param (
        [string]$url,
        [string]$outputPath
    )
    Write-Host "Downloading Citrix Workspace installer..."
    try {
        Invoke-WebRequest -Uri $url -OutFile $outputPath -UseBasicParsing
        Write-Host "Download completed successfully."
    } catch {
        Write-Host "An error occurred during the download: $_"
        exit 1
    }
}

# Check if the installer file exists
if (!(Test-Path -Path $installerPath)) {
    Write-Host "The installer file was not found at: $installerPath"
    Download-Installer -url $downloadUrl -outputPath $installerPath
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
