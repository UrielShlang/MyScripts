# Fetch the version of Citrix installed
$CitrixVersion = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" `
| Where-Object { $_.DisplayName -like "*Citrix*" } `
| Select-Object -ExpandProperty DisplayVersion -ErrorAction SilentlyContinue)

# Check if the version matches
if ($CitrixVersion -eq "24.2.1000.1016") {
    Write-Host "Citrix 24.2.1000.1016 Installed!"
	Exit 0  # Installed
} else {
    Exit 1  # Not Installed
}