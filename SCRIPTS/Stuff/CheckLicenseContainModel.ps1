#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string]$PolicyName = 'enterprisemgmt_policymanager-License-MDMPolicyAllowList',
    [string]$CheckValue = 'SecureBoot',
    [string]$CsvPath = "$env:TEMP\MDMPolicyAllowList_SecureBoot_Check.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ------------------------------------------------------------
# Native interop
# ------------------------------------------------------------
$source = @"
using System;
using System.Runtime.InteropServices;

namespace Win32Licensing
{
    public enum SLDATATYPE : uint
    {
        SL_DATA_NONE     = 0,
        SL_DATA_SZ       = 1,
        SL_DATA_BINARY   = 3,
        SL_DATA_DWORD    = 4,
        SL_DATA_MULTI_SZ = 7,
        SL_DATA_SUM      = 100
    }

    public static class NativeMethods
    {
        [DllImport("slc.dll", CharSet = CharSet.Unicode, SetLastError = false)]
        public static extern int SLGetWindowsInformation(
            string pwszValueName,
            out SLDATATYPE peDataType,
            out uint pcbValue,
            out IntPtr ppbValue
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr LocalFree(IntPtr hMem);
    }
}
"@

if (-not ('Win32Licensing.NativeMethods' -as [type])) {
    Add-Type -TypeDefinition $source -Language CSharp
}

# ------------------------------------------------------------
# Helpers
# ------------------------------------------------------------
function Convert-HResultToHex {
    param(
        [Parameter(Mandatory)]
        [int]$HResult
    )

    '0x{0:X8}' -f ([System.BitConverter]::ToUInt32([System.BitConverter]::GetBytes($HResult), 0))
}

function Convert-SLData {
    param(
        [Parameter(Mandatory)]
        [Win32Licensing.SLDATATYPE]$Type,

        [AllowNull()]
        [AllowEmptyCollection()]
        [byte[]]$Bytes
    )

    if ($null -eq $Bytes -or $Bytes.Count -eq 0) {
        return $null
    }

    switch ([uint32]$Type) {
        1 { # REG_SZ
            ([System.Text.Encoding]::Unicode.GetString($Bytes)).TrimEnd([char]0)
        }

        4 { # REG_DWORD
            if ($Bytes.Length -ge 4) {
                [BitConverter]::ToUInt32($Bytes, 0)
            }
            else {
                $null
            }
        }

        3 { # REG_BINARY
            (($Bytes | ForEach-Object { $_.ToString('X2') }) -join '')
        }

        7 { # REG_MULTI_SZ
            $raw = [System.Text.Encoding]::Unicode.GetString($Bytes).TrimEnd([char]0)
            $items = $raw -split [char]0 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            if ($items) { $items -join '; ' } else { $null }
        }

        100 { # SL_DATA_SUM
            if ($Bytes.Length -ge 4) {
                [BitConverter]::ToUInt32($Bytes, 0)
            }
            else {
                (($Bytes | ForEach-Object { $_.ToString('X2') }) -join '')
            }
        }

        default {
            (($Bytes | ForEach-Object { $_.ToString('X2') }) -join '')
        }
    }
}

function Invoke-SLGetWindowsInformation {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $dataType = [Win32Licensing.SLDATATYPE]::SL_DATA_NONE
    [uint32]$cb = 0
    $ptr = [IntPtr]::Zero

    try {
        $hr = [Win32Licensing.NativeMethods]::SLGetWindowsInformation(
            $Name,
            [ref]$dataType,
            [ref]$cb,
            [ref]$ptr
        )

        $hrHex = Convert-HResultToHex -HResult $hr

        if ($hr -ne 0) {
            return [PSCustomObject]@{
                Name       = $Name
                Found      = $false
                HResult    = $hr
                HResultHex = $hrHex
                DataType   = $null
                Size       = 0
                Value      = $null
            }
        }

        [byte[]]$bytes = $null
        if ($cb -gt 0 -and $ptr -ne [IntPtr]::Zero) {
            $bytes = New-Object byte[] ([int]$cb)
            [Runtime.InteropServices.Marshal]::Copy($ptr, $bytes, 0, [int]$cb)
        }

        $value = Convert-SLData -Type $dataType -Bytes $bytes

        [PSCustomObject]@{
            Name       = $Name
            Found      = $true
            HResult    = $hr
            HResultHex = $hrHex
            DataType   = $dataType.ToString()
            Size       = $cb
            Value      = $value
        }
    }
    finally {
        if ($ptr -ne [IntPtr]::Zero) {
            [void][Win32Licensing.NativeMethods]::LocalFree($ptr)
        }
    }
}

function Split-AllowListValue {
    param(
        [AllowNull()]
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    @(
        $Value -split '\|' |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

# ------------------------------------------------------------
# Main
# ------------------------------------------------------------
Write-Host "Checking policy: $PolicyName" -ForegroundColor Cyan

$result = Invoke-SLGetWindowsInformation -Name $PolicyName

if (-not $result.Found) {
    Write-Host "Policy lookup failed." -ForegroundColor Red
    $result | Format-List *

    $result | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Saved CSV to: $CsvPath" -ForegroundColor Yellow
    return
}

$rawValue = [string]$result.Value
$items = Split-AllowListValue -Value $rawValue
$containsValue = [bool]($items | Where-Object { $_ -ieq $CheckValue } | Select-Object -First 1)

$summary = [PSCustomObject]@{
    PolicyName      = $PolicyName
    Found           = $result.Found
    HResultHex      = $result.HResultHex
    DataType        = $result.DataType
    Size            = $result.Size
    ItemCount       = $items.Count
    CheckValue      = $CheckValue
    ContainsValue   = $containsValue
}

Write-Host ""
Write-Host "===== Summary =====" -ForegroundColor Cyan
$summary | Format-List

Write-Host ""
Write-Host "===== Matching Item Check =====" -ForegroundColor Cyan
if ($containsValue) {
    Write-Host "$CheckValue exists in the allow list." -ForegroundColor Green
}
else {
    Write-Host "$CheckValue does NOT exist in the allow list." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "===== All AllowList Items =====" -ForegroundColor Cyan
$items | Sort-Object | ForEach-Object {
    if ($_ -ieq $CheckValue) {
        Write-Host " * $_  <-- MATCH" -ForegroundColor Green
    }
    else {
        Write-Host " - $_"
    }
}

$export = @()

$export += [PSCustomObject]@{
    Section       = 'Summary'
    PolicyName    = $summary.PolicyName
    Found         = $summary.Found
    HResultHex    = $summary.HResultHex
    DataType      = $summary.DataType
    Size          = $summary.Size
    ItemCount     = $summary.ItemCount
    CheckValue    = $summary.CheckValue
    ContainsValue = $summary.ContainsValue
    Value         = $rawValue
}

$export += $items | ForEach-Object {
    [PSCustomObject]@{
        Section       = 'AllowListItem'
        PolicyName    = $PolicyName
        Found         = $result.Found
        HResultHex    = $result.HResultHex
        DataType      = $result.DataType
        Size          = $result.Size
        ItemCount     = $items.Count
        CheckValue    = $CheckValue
        ContainsValue = ($_ -ieq $CheckValue)
        Value         = $_
    }
}

$export | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Saved CSV to: $CsvPath" -ForegroundColor Green