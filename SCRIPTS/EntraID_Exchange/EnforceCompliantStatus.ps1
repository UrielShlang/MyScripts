Install-Module -Name aadinternals
import-module -Name aadinternals
Get-AADIntAccessTokenForAADGraph -UseMSAL -SaveToCache
$AzureDeviceIDHere="4c34b209-7468-4bf9-bdee-842c3a95f6c2"
Get-AADIntDeviceCompliance -deviceId $AzureDeviceIDHere
Set-AADIntDeviceCompliant -DeviceId $AzureDeviceIDHere -Compliant