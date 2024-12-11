# Get the current user's SID
##
##
$userLogin = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$sid = "HKEY_USERS\" + $userLogin

# Define the NotifyIconSettings path
$notifyIconSettingsPath = "$sid\Control Panel\NotifyIconSettings"

# Initialize a flag to track the detection result
$isCompliant = $true

# Check if the path exists
if (Test-Path -Path "Registry::$notifyIconSettingsPath") {
    Write-Host "Found path: $notifyIconSettingsPath" -ForegroundColor Green

    # Get all subkeys under NotifyIconSettings
    $childKeys = Get-ChildItem -Path "Registry::$notifyIconSettingsPath"

    foreach ($childKey in $childKeys) {
        $childKeyPath = $childKey.PSPath

        # Check if the "IsPromoted" DWORD exists and has a value of 1
        try {
            $isPromotedValue = Get-ItemProperty -Path $childKeyPath -Name "IsPromoted" -ErrorAction Stop
            if ($isPromotedValue.IsPromoted -ne 1) {
                Write-Host "Key $childKeyPath is non-compliant: 'IsPromoted' is not 1" -ForegroundColor Yellow
                $isCompliant = $false
            } else {
                Write-Host "Key $childKeyPath is compliant" -ForegroundColor Cyan
            }
        } catch {
            Write-Host "Key $childKeyPath is non-compliant: 'IsPromoted' not found" -ForegroundColor Red
            $isCompliant = $false
        }
    }
} else {
    Write-Host "Path not found: $notifyIconSettingsPath" -ForegroundColor Yellow
    $isCompliant = $false
}

# Return the compliance result
if ($isCompliant) {
    Write-Host "All keys are compliant" -ForegroundColor Green
    exit 0  # Exit with success code
} else {
    Write-Host "Non-compliance detected" -ForegroundColor Red
    exit 1  # Exit with failure code
}