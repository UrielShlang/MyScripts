# התקנת מודולים נדרשים
if (!(Get-Module -ListAvailable -Name Az.OperationalInsights)) {
    Install-Module -Name Az.OperationalInsights -Force -AllowClobber
}
if (!(Get-Module -ListAvailable -Name Microsoft.Graph.DeviceManagement)) {
    Install-Module -Name Microsoft.Graph.DeviceManagement -Force -AllowClobber
}
if (!(Get-Module -ListAvailable -Name Microsoft.Graph.Users)) {
    Install-Module -Name Microsoft.Graph.Users -Force -AllowClobber
}

$TenantID=""
$ResourceGroupName_Agent=""
$WorkspaceName_Agent=""
$ReportName = ""
$SaverLocation = "c:\temp"
$FullPathLocation = Join-Path $SaverLocation "$ReportName.csv"


function Get-UserAccountInfo {
    param(
        [string]$UserPrincipalName
    )
    
    if ([string]::IsNullOrEmpty($UserPrincipalName) -or $UserPrincipalName -eq "Not Found") {
        return @{
            AccountStatus = "Unknown"
            DisplayName = "N/A"
        }
    }
    
    try {
        $user = Get-MgUser -UserId $UserPrincipalName -Property "AccountEnabled,UserPrincipalName,DisplayName" -ErrorAction Stop
        
        $status = if ($user.AccountEnabled) { "Active" } else { "Blocked" }
        $displayName = if ([string]::IsNullOrEmpty($user.DisplayName)) { "N/A" } else { $user.DisplayName }
        
        return @{
            AccountStatus = $status
            DisplayName = $displayName
        }
    } catch {
        if ($_.Exception.Message -like "*Request_ResourceNotFound*" -or $_.Exception.Message -like "*does not exist*") {
            return @{
                AccountStatus = "User Not Found"
                DisplayName = "N/A"
            }
        } else {
            Write-Warning "Error checking user $UserPrincipalName : $($_.Exception.Message)"
            return @{
                AccountStatus = "Error"
                DisplayName = "N/A"
            }
        }
    }
}

function Get-UCClientReadinessReport {
    param(
        [string]$ResourceGroupName = $ResourceGroupName_Agent,
        [string]$WorkspaceName = $WorkspaceName_Agent,
        [string]$ReadinessStatus = "",
        [string]$IneligibilityReason = "",
        [datetime]$SnapShotTime = [datetime]::SpecifyKind((Get-Date).AddDays(-1).Date.AddHours(22), 'Utc'),
        [string]$OutputPath = "UCClientReadinessReport.csv",
        [switch]$IncludeUPN,
        [switch]$CheckAccountStatus
    )
  
    # בניית השאילתה
    $query = @"
let _ReadinessStatus="$ReadinessStatus";
let _IneligibilityReason="$IneligibilityReason";
let _SnapShotTime = datetime($($SnapShotTime.ToString('yyyy-MM-ddTHH:mm:ss')));
let UCClient_Info = UCClient | where TimeGenerated == _SnapShotTime;
UCClientReadinessStatus 
| join kind=leftouter (UCClient_Info) on AzureADDeviceId
| where TimeGenerated == _SnapShotTime
| where OSVersion has "Windows 10"
| where iff(_ReadinessStatus has "All", true, ReadinessStatus has _ReadinessStatus)
| where iff(_IneligibilityReason has "ALL", true, ReadinessReason has _IneligibilityReason)
| project DeviceName, AzureADDeviceId, OSVersion, ReadinessStatus, ReadinessReason, OSFeatureUpdateEOSTime
"@
    
    try {
        # הרצת שאילתה ב-Log Analytics
        Write-Host "Executing Log Analytics query..." -ForegroundColor Yellow
        # חיבור לשירותים
        Disconnect-AzAccount
        Connect-AzAccount -Tenant $TenantID -ErrorAction Stop

        $workspace = Get-AzOperationalInsightsWorkspace -ResourceGroupName $ResourceGroupName -Name $WorkspaceName
        $results = Invoke-AzOperationalInsightsQuery -WorkspaceId $workspace.CustomerId -Query $query
        
        if ($results.Results.Count -eq 0) {
            Write-Warning "No results found."
            return
        }
        
        Write-Host "Found $($results.Results.Count) devices" -ForegroundColor Green
        
        # הוספת UPN אם נדרש
        if ($IncludeUPN -or $CheckAccountStatus) {
            Write-Host "Fetching UPN data from Microsoft Graph..." -ForegroundColor Yellow
            
            # הרשאות נדרשות
            $scopes = @("Device.Read.All", "DeviceManagementManagedDevices.Read.All")
            if ($CheckAccountStatus) {
                $scopes += "User.Read.All"
            }
            
            Connect-MgGraph -Scopes $scopes -ErrorAction Stop

            # יצירת hash table לחיפוש מהיר של devices
            $deviceToUPN = @{}
            
            # קבלת נתונים מ-Intune Managed Devices (מכיל UPN ישירות)
            try {
                $intuneDevices = Get-MgDeviceManagementManagedDevice -All | Where-Object { $_.DeviceName -and $_.UserPrincipalName }
                foreach ($device in $intuneDevices) {
                    if (-not $deviceToUPN.ContainsKey($device.DeviceName)) {
                        Write-Host $device.UserPrincipalName
                        $deviceToUPN[$device.DeviceName] = $device.UserPrincipalName
                    }
                }
                Write-Host "Found UPN for $($deviceToUPN.Count) devices from Intune" -ForegroundColor Green
            } catch {
                Write-Warning "Could not access Intune managed devices. Trying alternative method..."
            }
            
            # הוספת UPN לתוצאות
            $enrichedResults = foreach ($result in $results.Results) {
                $upn = if ($deviceToUPN.ContainsKey($result.DeviceName)) { 
                    Write-Host $deviceToUPN[$result.DeviceName] 
                    $deviceToUPN[$result.DeviceName] 
                } else { 
                    "Not Found" 
                }
                
                # בדיקת סטטוס החשבון אם נדרש
                $userInfo = if ($CheckAccountStatus) {
                    Write-Host "Checking account status for: $upn" -ForegroundColor Yellow
                    Get-UserAccountInfo -UserPrincipalName $upn
                } else {
                    @{
                        AccountStatus = $null
                        DisplayName = $null
                    }
                }
                
                # החזרת אובייקט עם UPN, Display Name ו-Account Status
                $resultObject = [PSCustomObject]@{
                    DeviceName = $result.DeviceName
                    AzureADDeviceId = $result.AzureADDeviceId
                    OSVersion = $result.OSVersion
                    ReadinessStatus = $result.ReadinessStatus
                    ReadinessReason = $result.ReadinessReason
                    OSFeatureUpdateEOSTime = $result.OSFeatureUpdateEOSTime
                }
                
                if ($IncludeUPN) {
                    $resultObject | Add-Member -MemberType NoteProperty -Name "UPN" -Value $upn
                }
                
                if ($CheckAccountStatus) {
                    $resultObject | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $userInfo.DisplayName
                    $resultObject | Add-Member -MemberType NoteProperty -Name "AccountStatus" -Value $userInfo.AccountStatus
                }
                
                $resultObject
            }
            
            $finalResults = $enrichedResults
            $enrichedResults | Out-GridView
        } else {
            $finalResults = $results.Results
        }
        
        # שמירה וייצוא
        $finalResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8BOM
        Write-Host "Report saved to: $OutputPath" -ForegroundColor Green
        
        # הצגת סיכום
        Write-Host "`nSummary:" -ForegroundColor Cyan
        $finalResults | Group-Object ReadinessStatus | ForEach-Object {
            Write-Host "  $($_.Name): $($_.Count) devices" -ForegroundColor White
        }
        
        # סיכום נוסף לסטטוס חשבונות אם נבדק
        if ($CheckAccountStatus) {
            Write-Host "`nAccount Status Summary:" -ForegroundColor Cyan
            $finalResults | Group-Object AccountStatus | ForEach-Object {
                Write-Host "  $($_.Name): $($_.Count) accounts" -ForegroundColor White
            }
        }
        
        return $finalResults
        
    } catch {
        Write-Error "Error: $($_.Exception.Message)"
    } finally {
        if ($IncludeUPN -or $CheckAccountStatus) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue
        }
    }
}

# דוגמאות שימוש:
# Get-UCClientReadinessReport
# Get-UCClientReadinessReport -IncludeUPN
 Get-UCClientReadinessReport -IncludeUPN -CheckAccountStatus -OutputPath $FullPathLocation
# Get-UCClientReadinessReport -ReadinessStatus "Ready" -IncludeUPN -CheckAccountStatus -OutputPath "C:\Reports\ReadyDevices.csv"

# הדוח יכלול את העמודות הבאות (כשמופעלות האפשרויות המתאימות):
# - DeviceName: שם המכשיר
# - AzureADDeviceId: מזהה המכשיר ב-Azure AD  
# - OSVersion: גרסת מערכת הפעלה
# - ReadinessStatus: סטטוס מוכנות העדכון
# - ReadinessReason: סיבת אי-מוכנות (אם רלוונטי)
# - OSFeatureUpdateEOSTime: זמן סיום תמיכה בגרסה
# - UPN: שם המשתמש (אם IncludeUPN מופעל)
# - DisplayName: השם המלא של המשתמש (אם CheckAccountStatus מופעל)
# - AccountStatus: סטטוס החשבון - Active/Blocked/User Not Found/Error (אם CheckAccountStatus מופעל)