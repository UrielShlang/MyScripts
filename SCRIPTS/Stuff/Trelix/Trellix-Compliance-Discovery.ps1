<#
    Trellix-Compliance-Discovery.ps1
    Intune custom compliance discovery script.

    Built from real discovery output: this estate runs Trellix Agent + DLP Endpoint
    + DXL, with Cortex XDR as the actual antivirus. There is NO ENS Threat Prevention,
    so no AV/DAT/RTP checks are performed against Trellix.

    Intune upload settings:
      Run this script using the logged on credentials : No
      Enforce script signature check                  : No (unless you sign it)
      Run script in 64 bit PowerShell Host            : Yes   <-- important

    Output: single-line compressed JSON. Keys must match SettingName in the JSON file.
#>

$ErrorActionPreference = 'SilentlyContinue'

# ---------- helpers ----------
$uninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$installed = Get-ItemProperty $uninstallPaths | Where-Object { $_.DisplayName }

function Get-ProductVersion {
    param([string]$NamePattern)
    $hit = $installed |
        Where-Object { $_.DisplayName -match $NamePattern -and $_.DisplayVersion } |
        Sort-Object { [version]($_.DisplayVersion -replace '[^0-9.]', '') } -Descending |
        Select-Object -First 1
    if ($hit) {
        # normalise to 4-part version so Intune's Version comparison never chokes
        $v = ($hit.DisplayVersion -replace '[^0-9.]', '').Trim('.')
        $parts = $v.Split('.') | Where-Object { $_ -ne '' }
        while ($parts.Count -lt 4) { $parts += '0' }
        return ($parts[0..3] -join '.')
    }
    return '0.0.0.0'
}

function Get-ServiceState {
    param([string]$Name)
    $svc = Get-Service -Name $Name
    if (-not $svc) { return 'Not Installed' }
    if ($svc.Status -eq 'Running') { return 'Running' }
    return [string]$svc.Status
}

# ---------- Trellix components ----------
$trellixAgentVersion = Get-ProductVersion 'Trellix Agent|McAfee Agent'
$trellixDlpVersion   = Get-ProductVersion 'Trellix Data Loss Prevention|McAfee DLP Endpoint'
$trellixDxlVersion   = Get-ProductVersion 'Data Exchange Layer'

$trellixAgentService = Get-ServiceState 'masvc'
$trellixAgentCommon  = Get-ServiceState 'macmnsvc'
$trellixDlpService   = Get-ServiceState 'TrellixDLPAgentService'

# ---------- actual antivirus (Cortex XDR) ----------
# Trellix is not the AV on these devices; keep AV assurance anchored to what is.
$avExpected = 'Cortex XDR'
$avRtp  = 'Error: No Antivirus product found'
$avDefs = 'Error: No Antivirus product found'
$avName = 'Error: No Antivirus product found'

$av = Get-CimInstance -Namespace 'root\SecurityCenter2' -Class AntiVirusProduct |
      Where-Object { $_.displayName -like "*$avExpected*" } | Select-Object -First 1

if ($av) {
    $avName = $avExpected
    $hex = [Convert]::ToString($av.productState, 16).PadLeft(6, '0')
    $avRtp = switch ($hex.Substring(2, 2)) {
        '00' { 'Off' } '01' { 'Expired' } '10' { 'On' } '11' { 'Snoozed' } default { 'Unknown' }
    }
    $avDefs = switch ($hex.Substring(4, 2)) {
        '00' { 'Up to Date' } '10' { 'Out of Date' } default { 'Unknown' }
    }
}

# ---------- emit ----------
$result = [ordered]@{
    TrellixAgentVersion         = $trellixAgentVersion
    TrellixAgentService         = $trellixAgentService
    TrellixAgentCommonService   = $trellixAgentCommon
    TrellixDLPVersion           = $trellixDlpVersion
    TrellixDLPService           = $trellixDlpService
    TrellixDXLVersion           = $trellixDxlVersion
    AntivirusProduct            = $avName
    AntivirusRealTimeProtection = $avRtp
    AntivirusDefinitions        = $avDefs
}

return ([PSCustomObject]$result | ConvertTo-Json -Compress)
