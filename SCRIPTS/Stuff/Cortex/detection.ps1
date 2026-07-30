Function Check-Cortex {
$CortexTimeStamp=Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct | where displayName -Like "Cortex XDR*" | select timestamp 
#$CortexTimeStamp.timestamp

#[Datetime]$CortexTimeStamp.timestamp
$exe = (Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntiVirusProduct |
        Where-Object displayName -like 'Cortex XDR*').pathToSignedProductExe
(Get-Item $exe).VersionInfo.ProductVersion


$date1=Get-Date

if (([string]::IsNullOrEmpty($CortexTimeStamp)))
{
    #[System.Windows.MessageBox]::Show('null')
    Return @{Name = $false} | ConvertTo-Json -Compress
}
if ([Datetime]$CortexTimeStamp.timestamp -gt $date1.AddDays(-14))
    {
    return @{Name = $true} | ConvertTo-Json -Compress
    # [System.Windows.MessageBox]::Show('OK')
    }
    else
    {
    Return @{Name = $false} | ConvertTo-Json -Compress
    # [System.Windows.MessageBox]::Show('no OK')
    }

}
Check-Cortex
