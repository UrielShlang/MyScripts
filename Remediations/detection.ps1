# Get all subkeys under HKEY_USERS
$usersSubkeys = Get-ChildItem -Path "Registry::HKEY_USERS"

foreach ($userKey in $usersSubkeys) {
    $sid = $userKey.Name
    $notifyIconSettingsPath = "$sid\Control Panel\NotifyIconSettings"

    # Check if the path exists
    if (Test-Path -Path "Registry::$notifyIconSettingsPath") {
        Write-Host "Found path: $notifyIconSettingsPath" -ForegroundColor Green

        # Get all subkeys under NotifyIconSettings
        $childKeys = Get-ChildItem -Path "Registry::$notifyIconSettingsPath"

        foreach ($childKey in $childKeys) {
            $childKeyPath = $childKey.PSPath
            $isPromoted = Get-ItemProperty -Path $childKeyPath -Name "IsPromoted" -ErrorAction SilentlyContinue

            # Check if IsPromoted exists and its value is 1
            if ($isPromoted -and $isPromoted.IsPromoted -eq 1) {
                Write-Host "'IsPromoted' exists and has value 1 at $childKeyPath" -ForegroundColor Green
            } else {
                Write-Host "'IsPromoted' is either missing or not set to 1 at $childKeyPath" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "Path not found: $notifyIconSettingsPath" -ForegroundColor Yellow
    }
}
