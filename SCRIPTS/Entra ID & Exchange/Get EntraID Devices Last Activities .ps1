# Azure Automation Runbook - Non-Compliant Devices Report
# אימות באמצעות Managed Identity או App Registration
try {
    Write-Output "Connecting to Microsoft Graph..."
    
    # ניסיון עם Managed Identity
    try {
        Connect-MgGraph -Identity
        Write-Output "Successfully connected to Microsoft Graph using Managed Identity"
    }
    catch {
        Write-Output "Managed Identity failed, trying with stored credentials..."
        
        # כאן אפשר להוסיף חיבור עם App Registration אם יש
        # $tenantId = Get-AutomationVariable -Name 'TenantId'
        # $clientId = Get-AutomationVariable -Name 'ClientId'
        # $clientSecret = Get-AutomationVariable -Name 'ClientSecret'
        
        # אם אין אוטומציה משתנים, נסה שוב עם Identity
        Connect-MgGraph -Identity
        Write-Output "Successfully connected to Microsoft Graph"
    }
}
catch {
    Write-Error "Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    Write-Output "Make sure the Managed Identity has the required permissions:"
    Write-Output "- DeviceManagementConfiguration.Read.All"
    Write-Output "- DeviceManagementManagedDevices.Read.All"
    exit 1
}

# קבלת כל המכשירים הלא תואמים
try {
    Write-Output "Retrieving non-compliant devices..."
    $nonCompliantDevices = Get-MgDeviceManagementManagedDevice -Filter "complianceState eq 'noncompliant'" -All
    Write-Output "Found $($nonCompliantDevices.Count) non-compliant devices"
}
catch {
    Write-Error "Failed to retrieve non-compliant devices: $($_.Exception.Message)"
    exit 1
}

$deviceComplianceData = @{}

foreach ($device in $nonCompliantDevices) {
    Write-Output "Processing device: $($device.deviceName) - $($device.Id)"
    
    # יצירת מפתח ייחודי למכשיר
    $deviceKey = "$($device.deviceName)_$($device.Id)"
    
    # אתחול נתוני המכשיר אם לא קיים
    if (-not $deviceComplianceData.ContainsKey($deviceKey)) {
        $deviceComplianceData[$deviceKey] = @{
            DeviceName = $device.deviceName
            DeviceId = $device.Id
            UserPrincipalName = $device.userPrincipalName
            NonCompliantSettings = @()
            NonCompliantPolicies = @()
        }
    }
    
    try {
        # קבלת מצב הפוליסיות עבור המכשיר הספציפי
        $policyStates = Invoke-MgGraphRequest -Uri "/beta/deviceManagement/managedDevices/$($device.Id)/deviceCompliancePolicyStates" -Method GET
        
        foreach ($policyState in $policyStates.value) {
            if ($policyState.state -eq "nonCompliant") {
                # הוספת שם הפוליסי הלא תואם
                if ($deviceComplianceData[$deviceKey].NonCompliantPolicies -notcontains $policyState.displayName) {
                    $deviceComplianceData[$deviceKey].NonCompliantPolicies += $policyState.displayName
                }
                
                # קבלת הפרטים הספציפיים של המכשיר עבור הפוליסי
                try {
                    $deviceStates = Invoke-MgGraphRequest -Uri "/beta/deviceManagement/managedDevices/$($device.Id)/deviceCompliancePolicyStates/$($policyState.id)/settingStates" -Method GET
                    
                    foreach ($settingState in $deviceStates.value) {
                        if ($settingState.state -eq "nonCompliant") {
                            # שימוש רק בערך של ה-setting
                            $settingDetail = $settingState.setting
                            
                            # אם אין ערך ב-setting, שימוש ב-settingName
                            if (-not $settingDetail -or $settingDetail -eq "") {
                                $settingDetail = $settingState.settingName
                            }
                            
                            # הוספת ההגדרה אם לא קיימת כבר
                            if ($deviceComplianceData[$deviceKey].NonCompliantSettings -notcontains $settingDetail) {
                                $deviceComplianceData[$deviceKey].NonCompliantSettings += $settingDetail
                            }
                        }
                    }
                }
                catch {
                    Write-Warning "Could not get setting states for policy $($policyState.displayName) on device $($device.deviceName): $($_.Exception.Message)"
                    # הוספת הודעת שגיאה כללית
                    $generalPolicyError = "Unable to retrieve settings for policy: $($policyState.displayName)"
                    if ($deviceComplianceData[$deviceKey].NonCompliantSettings -notcontains $generalPolicyError) {
                        $deviceComplianceData[$deviceKey].NonCompliantSettings += $generalPolicyError
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Could not get policy states for device $($device.deviceName): $($_.Exception.Message)"
    }
}

# יצירת הדוח הסופי
$ConsolidatedReport = @()

foreach ($deviceKey in $deviceComplianceData.Keys) {
    $deviceData = $deviceComplianceData[$deviceKey]
    
    $ConsolidatedReport += [PSCustomObject]@{
        DeviceName = $deviceData.DeviceName
        DeviceId = $deviceData.DeviceId
        UserPrincipalName = $deviceData.UserPrincipalName
        NonCompliantPoliciesCount = $deviceData.NonCompliantPolicies.Count
        NonCompliantPolicies = ($deviceData.NonCompliantPolicies -join "; ")
        NonCompliantSettingsCount = $deviceData.NonCompliantSettings.Count
        NonCompliantSettings = ($deviceData.NonCompliantSettings -join "; ")
    }
}

# הצגת התוצאות
if ($ConsolidatedReport.Count -gt 0) {
    # הצגה כטבלה ב-Output (במקום Out-GridView שלא זמין ב-Azure Automation)
    Write-Output "=== CONSOLIDATED NON-COMPLIANT DEVICES REPORT ==="
    Write-Output "Total non-compliant devices: $($ConsolidatedReport.Count)"
    Write-Output "Total non-compliant settings across all devices: $(($ConsolidatedReport | Measure-Object -Property NonCompliantSettingsCount -Sum).Sum)"
    Write-Output ""
    
    # הצגת הדוח
    $ConsolidatedReport | Format-Table -AutoSize | Out-String | Write-Output
    
    # יצירת תוכן CSV לשליחה או לשמירה
    $csvContent = $ConsolidatedReport | ConvertTo-Csv -NoTypeInformation
    
    # הצגת תוכן CSV ב-Output
    Write-Output "=== CSV FORMAT ==="
    $csvContent | Write-Output
    
    # אופציה לשמירה ב-Storage Account (אם מוגדר)
    # אפשר להוסיף קוד כאן לשליחת הדוח למייל או לשמירה ב-Blob Storage
    
    Write-Output "=== REPORT COMPLETED SUCCESSFULLY ==="
    
} else {
    Write-Output "No detailed compliance information found. This might be because:"
    Write-Output "1. The devices don't have detailed compliance policy states available"
    Write-Output "2. The API endpoints might require different permissions"
    Write-Output "3. The compliance policies might not be configured to report detailed states"
}

# ניתוק מ-Graph
try {
    Disconnect-MgGraph
    Write-Output "Disconnected from Microsoft Graph"
}
catch {
    Write-Warning "Could not disconnect from Microsoft Graph: $($_.Exception.Message)"
}