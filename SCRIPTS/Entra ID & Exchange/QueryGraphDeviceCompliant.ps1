Connect-MgGraph -Scopes DeviceManagementConfiguration.Read.All,DeviceManagementManagedDevices.Read.All

# קבלת כל המכשירים הלא תואמים
$nonCompliantDevices = Get-MgDeviceManagementManagedDevice -Filter "complianceState eq 'noncompliant'" -All

$Detailed = @()

foreach ($device in $nonCompliantDevices) {
    Write-Host "Processing device: $($device.deviceName) - $($device.Id)"
    try {
        #$device.Id="f4cd497d-892e-4f5f-a719-2157f0e78e34"
        # קבלת מצב הפוליסיות עבור המכשיר הספציפי
        $policyStates = Invoke-MgGraphRequest -Uri "/beta/deviceManagement/managedDevices/$($device.Id)/deviceCompliancePolicyStates" -Method GET
        
        foreach ($policyState in $policyStates.value) {
            if ($policyState.state -in @("nonCompliant", "error","unknown")) {
                # קבלת הפרטים הספציפיים של המכשיר עבור הפוליסי
                try {
                    #$policyState.id="2b521800-b0df-4959-ae47-0bd5d01f7a47"
                    $deviceStates = Invoke-MgGraphRequest -Uri "/beta/deviceManagement/managedDevices/$($device.Id)/deviceCompliancePolicyStates/$($policyState.id)/settingStates" -Method GET
                    foreach ($settingState in $deviceStates.value) {
                        #Write-Host $settingState.state
                        if ($settingState.state -in @("nonCompliant", "error","unknown")) {
                            $Detailed += [PSCustomObject]@{
                                DeviceName      = $device.deviceName
                                DeviceId        = $device.Id
                                UserPrincipalName = $device.userPrincipalName
                                PolicyName      = $policyState.displayName
                                PolicyId        = $policyState.id
                                Setting         = $settingState.setting
                                SettingName     = $settingState.
                                Description = $settingState.errorDescription
                                ErrorCode       = $settingState.errorCode
                                State           = $settingState.state
                                DeviceLastSyncDateTime    = $device.LastSyncDateTime                    
                            }
                        }
                    }
                   #Write-Host $Detailed 
                }
                catch {
                    Write-Warning "Could not get setting states for policy $($policyState.displayName) on device $($device.deviceName): $($_.Exception.Message)"
                }
            }
        }
    }
    catch {
        Write-Warning "Could not get policy states for device $($device.deviceName): $($_.Exception.Message)"
    }
}

# הצגת התוצאות
if ($Detailed.Count -gt 0) {
    $Detailed | Out-GridView
    #$Detailed | Format-Table -AutoSize | Out-String | Write-Output
    Write-Host "Total non-compliant settings found: $($Detailed.Count)"
} else {
    Write-Host "No detailed compliance information found. This might be because:"
    Write-Host "1. The devices don't have detailed compliance policy states available"
    Write-Host "2. The API endpoints might require different permissions"
    Write-Host "3. The compliance policies might not be configured to report detailed states"
}