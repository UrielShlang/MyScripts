$CitrixVersion = (Get-ItemProperty -Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -like "*Citrix Workspace 2402*" }).DisplayVersion
if ($CitrixVersion -eq "24.2.1000.1016") {
    Exit 0  # Installed
} else {
    Exit 1  # Not Installed
}
