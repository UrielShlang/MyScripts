<#
    Trellix-Compliance-Discovery.ps1
    Intune custom compliance discovery script - Trellix components only.

    Scope: Trellix Agent, Trellix DLP Endpoint, Trellix Data Exchange Layer.
    No antivirus checks: this estate has no ENS Threat Prevention installed,
    and Agent/DLP/DXL do not register with Windows Security Center by design.

    Intune upload settings:
      Run this script using the logged on credentials : No
      Enforce script signature check                  : No (unless you sign it)
      Run script in 64 bit PowerShell Host            : Yes   <-- required

    Output: single-line compressed JSON. Keys must match SettingName in the JSON file.
#>

$ErrorActionPreference = 'SilentlyContinue'

# ---------- helpers ----------
$uninstallPaths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
$installed = Get-ItemProperty $uninstallPaths | Where-Object { $_.DisplayName }

function ConvertTo-FourPart {
    param([string]$Raw)
    $v = ($Raw -replace '[^0-9.]', '').Trim('.')
    if ([string]::IsNullOrWhiteSpace($v)) { return '0.0.0.0' }
    $parts = @($v.Split('.') | Where-Object { $_ -ne '' })
    while ($parts.Count -lt 4) { $parts += '0' }
    return ($parts[0..3] -join '.')
}

function Get-ProductVersion {
    param([string]$NamePattern)
    $versions = $installed |
        Where-Object { $_.DisplayName -match $NamePattern -and $_.DisplayVersion } |
        ForEach-Object { ConvertTo-FourPart $_.DisplayVersion }
    if (-not $versions) { return '0.0.0.0' }
    # multiple uninstall entries are common (Agent/DXL register twice) - take highest
    return ([string](($versions | ForEach-Object { [version]$_ } | Sort-Object -Descending)[0]))
}

function Get-ServiceState {
    param([string]$Name)
    $svc = Get-Service -Name $Name
    if (-not $svc) { return 'Not Installed' }
    if ($svc.Status -eq 'Running') { return 'Running' }
    return [string]$svc.Status
}

# ---------- Trellix components ----------
$result = [ordered]@{
    TrellixAgentVersion       = Get-ProductVersion 'Trellix Agent|McAfee Agent'
    TrellixAgentService       = Get-ServiceState   'masvc'
    TrellixAgentCommonService = Get-ServiceState   'macmnsvc'
    TrellixDLPVersion         = Get-ProductVersion 'Trellix Data Loss Prevention|McAfee DLP Endpoint'
    TrellixDLPService         = Get-ServiceState   'TrellixDLPAgentService'
    TrellixDXLVersion         = Get-ProductVersion 'Data Exchange Layer'
}

return ([PSCustomObject]$result | ConvertTo-Json -Compress)
