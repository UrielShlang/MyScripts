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
    $Detailed
    $Detailed | Format-Table -AutoSize | Out-String | Write-Output
    Write-Host "Total non-compliant settings found: $($Detailed.Count)"
} else {
    Write-Host "No detailed compliance information found. This might be because:"
    Write-Host "1. The devices don't have detailed compliance policy states available"
    Write-Host "2. The API endpoints might require different permissions"
    Write-Host "3. The compliance policies might not be configured to report detailed states"
}


# ניתוק מ-Graph
try {
    Disconnect-MgGraph
    Write-Output "Disconnected from Microsoft Graph"
}
catch {
    Write-Warning "Could not disconnect from Microsoft Graph: $($_.Exception.Message)"
}

### Step 0: Set variables required for the rest of the script.

# information needed to send data to the DCR endpoint
$endpoint_uri = "https://dce-compliantdevices-b9s1.ukwest-1.ingest.monitor.azure.com" #Logs ingestion URI for the DCR
$dcrImmutableId = "dcr-4e9a928251034b97a4ad206fee615a7d" #the immutableId property of the DCR object
$streamName = "Custom-TBL_CompliantDevices_CL" #name of the stream in the DCR that represents the destination table

### Step 1: Obtain a bearer token using Managed Identity.

Write-Output "Try to Send to Log Analytics"

# Optional: Set client_id for user-assigned managed identity (leave empty for system-assigned)
$clientId = "" # Set this to your user-assigned managed identity client ID if needed

try {
    $resourceUri = "https://monitor.azure.com/"
    
    # Build the token request URI
    $tokenAuthURI = $env:IDENTITY_ENDPOINT + "?resource=" + $resourceUri + "&api-version=2019-08-01"
    
    # Add client_id parameter if using user-assigned managed identity
    if ($clientId) {
        $tokenAuthURI += "&client_id=" + $clientId
    }
    
    # Request token using managed identity
    $tokenResponse = Invoke-RestMethod -Method Get -Headers @{"X-IDENTITY-HEADER"="$env:IDENTITY_HEADER"} -Uri $tokenAuthURI
    $bearerToken = $tokenResponse.access_token
    
    Write-Host "Successfully obtained bearer token using Managed Identity"
}
catch {
    Write-Error "Failed to obtain bearer token: $($_.Exception.Message)"
    exit 1
}

### Step 2: Create some sample data.

#$currentTime = Get-Date ([datetime]::UtcNow) -Format O

# יצוא עם הוספת timestamp לכל רשומה (מומלץ עבור Log Analytics)
$staticData = $Detailed | ForEach-Object {
    $_ | Add-Member -NotePropertyName "TimeGenerated" -NotePropertyValue (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss.fffZ") -Force
    $_
} | ConvertTo-Json -Depth 5

### Step 3: Send the data to the Log Analytics workspace.

$body = $staticData;
$headers = @{"Authorization"="Bearer $bearerToken";"Content-Type"="application/json"};
$uri = "$endpoint_uri/dataCollectionRules/$dcrImmutableId/streams/$($streamName)?api-version=2023-01-01"

try {
    $uploadResponse = Invoke-RestMethod -Uri $uri -Method "Post" -Body $body -Headers $headers
    Write-Host "Data sent successfully to Log Analytics"
    $uploadResponse
}
catch {
    Write-Error "Failed to send data to Log Analytics: $($_.Exception.Message)"
    exit 1
}