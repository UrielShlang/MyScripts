<#
    Discover-TrellixAnchors.ps1
    Run as SYSTEM or elevated admin on a KNOWN-GOOD machine (Trellix installed,
    protection on, definitions current). Output = the anchors you can safely use
    in an Intune custom compliance discovery script.

    Usage:  powershell -ExecutionPolicy Bypass -File .\Discover-TrellixAnchors.ps1
    Tip:    add  > C:\Temp\trellix-anchors.txt  and send me the file.
#>

function Section($t) { Write-Output "`n===== $t =====" }

Section "OS"
$os = Get-CimInstance Win32_OperatingSystem
Write-Output ("Caption      : {0}" -f $os.Caption)
Write-Output ("Version      : {0}" -f $os.Version)
Write-Output ("ProductType  : {0}  (1=Workstation, 2=DC, 3=Server)" -f $os.ProductType)

Section "SecurityCenter2 (expected empty on Server)"
try {
    Get-CimInstance -Namespace 'root\SecurityCenter2' -Class AntiVirusProduct -ErrorAction Stop |
        Select-Object displayName, productState, pathToSignedProductExe |
        Format-List
} catch {
    Write-Output "NOT AVAILABLE: $($_.Exception.Message)"
}

Section "Services matching mfe* / ma*svc / trellix / xagt"
Get-Service | Where-Object {
    $_.Name -match '^(mfe|ma[cs]|mms)' -or
    $_.Name -match 'trellix|xagt|mcafee' -or
    $_.DisplayName -match 'Trellix|McAfee|FireEye'
} | Select-Object Name, DisplayName, Status, StartType | Format-Table -AutoSize

Section "Installed products (uninstall keys)"
$paths = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
)
Get-ItemProperty $paths -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match 'Trellix|McAfee|FireEye' } |
    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
    Sort-Object DisplayName | Format-Table -AutoSize

Section "Registry tree: vendor roots (2 levels)"
foreach ($root in 'HKLM:\SOFTWARE\McAfee','HKLM:\SOFTWARE\Trellix',
                  'HKLM:\SOFTWARE\WOW6432Node\McAfee','HKLM:\SOFTWARE\WOW6432Node\Trellix') {
    if (Test-Path $root) {
        Write-Output "`n-- $root"
        Get-ChildItem $root -Recurse -Depth 2 -ErrorAction SilentlyContinue |
            Select-Object -ExpandProperty Name
    }
}

Section "Values under likely AV/DAT keys"
$candidates = @(
    'HKLM:\SOFTWARE\McAfee\Endpoint\AV',
    'HKLM:\SOFTWARE\McAfee\AVSolution\DS\DS',
    'HKLM:\SOFTWARE\McAfee\AVSolution\McTaskManager\OAS',
    'HKLM:\SOFTWARE\Wow6432Node\McAfee\AVEngine',
    'HKLM:\SOFTWARE\McAfee\Agent',
    'HKLM:\SOFTWARE\Wow6432Node\Network Associates\ePolicy Orchestrator\Application Plugins'
)
foreach ($key in $candidates) {
    if (Test-Path $key) {
        Write-Output "`n-- $key"
        (Get-ItemProperty $key) | Format-List *
    } else {
        Write-Output "`n-- $key  [absent]"
    }
}

Section "Key binaries + file versions"
$files = @(
    "$env:ProgramFiles\McAfee\Endpoint Security\Threat Prevention\mfetp.exe",
    "$env:ProgramFiles\McAfee\Endpoint Security\Endpoint Security Platform\mfeesp.exe",
    "$env:ProgramFiles\McAfee\Agent\masvc.exe",
    "$env:ProgramFiles\Trellix\Endpoint Security\Threat Prevention\mfetp.exe",
    "$env:ProgramFiles(x86)\McAfee\Endpoint Security\Threat Prevention\mfetp.exe"
)
foreach ($f in $files) {
    if (Test-Path $f) {
        $v = (Get-Item $f).VersionInfo
        Write-Output ("{0}`n    ProductVersion={1}  FileVersion={2}" -f $f, $v.ProductVersion, $v.FileVersion)
    }
}

Section "DAT / content files (newest 5 under vendor dirs)"
foreach ($d in "$env:ProgramData\McAfee","$env:ProgramData\Trellix") {
    if (Test-Path $d) {
        Write-Output "`n-- $d"
        Get-ChildItem $d -Recurse -Include '*.dat','avvclean.dat','avvnames.dat','avvscan.dat' -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 5 FullName, LastWriteTime |
            Format-Table -AutoSize
    }
}

Write-Output "`n===== DONE ====="
