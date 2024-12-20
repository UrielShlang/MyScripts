##
##
##
$userLogin=[System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$sid="HKEY_USERS\"+$userLogin

    #$sid = $userKey.Name
    $notifyIconSettingsPath = "$sid\Control Panel\NotifyIconSettings"

    # Check if the path exists
    if (Test-Path -Path "Registry::$notifyIconSettingsPath") {
        Write-Host "Found path: $notifyIconSettingsPath" -ForegroundColor Green

        # Get all subkeys under NotifyIconSettings
        $childKeys = Get-ChildItem -Path "Registry::$notifyIconSettingsPath"

        foreach ($childKey in $childKeys) {
            $childKeyPath = $childKey.PSPath

            # Add the "IsPromoted" DWORD with value 1
            try {
                New-ItemProperty -Path $childKeyPath -Name "IsPromoted" -PropertyType DWORD -Value 1 -Force
                Write-Host "Successfully added 'IsPromoted' to $childKeyPath" -ForegroundColor Cyan
            } catch {
                Write-Host "Failed to add 'IsPromoted' to " -ForegroundColor Red
            }
        }
    } else {
        Write-Host "Path not found: $notifyIconSettingsPath" -ForegroundColor Yellow
    }
