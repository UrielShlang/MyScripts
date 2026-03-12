#Requires -RunAsAdministrator
[CmdletBinding()]
param(
    [string]$PolicyName = 'enterprisemgmt_policymanager-License-MDMPolicyAllowList',
    [string]$CheckValue = 'SecureBoot'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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
        1 { ([System.Text.Encoding]::Unicode.GetString($Bytes)).TrimEnd([char]0) }
        4 { if ($Bytes.Length -ge 4) { [BitConverter]::ToUInt32($Bytes, 0) } else { $null } }
        3 { (($Bytes | ForEach-Object { $_.ToString('X2') }) -join '') }
        7 {
            $raw = [System.Text.Encoding]::Unicode.GetString($Bytes).TrimEnd([char]0)
            $items = $raw -split [char]0 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            if ($items) { $items -join '; ' } else { $null }
        }
        100 { if ($Bytes.Length -ge 4) { [BitConverter]::ToUInt32($Bytes, 0) } else { (($Bytes | ForEach-Object { $_.ToString('X2') }) -join '') } }
        default { (($Bytes | ForEach-Object { $_.ToString('X2') }) -join '') }
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

        if ($hr -ne 0) {
            return [PSCustomObject]@{
                Name       = $Name
                Found      = $false
                HResult    = $hr
                HResultHex = ('0x{0:X8}' -f ([System.BitConverter]::ToUInt32([System.BitConverter]::GetBytes($hr), 0)))
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

        [PSCustomObject]@{
            Name       = $Name
            Found      = $true
            HResult    = $hr
            HResultHex = ('0x{0:X8}' -f ([System.BitConverter]::ToUInt32([System.BitConverter]::GetBytes($hr), 0)))
            DataType   = $dataType.ToString()
            Size       = $cb
            Value      = Convert-SLData -Type $dataType -Bytes $bytes
        }
    }
    finally {
        if ($ptr -ne [IntPtr]::Zero) {
            [void][Win32Licensing.NativeMethods]::LocalFree($ptr)
        }
    }
}

function Split-AllowListValue {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return @()
    }

    @(
        $Value -split '\|' |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

$result = Invoke-SLGetWindowsInformation -Name $PolicyName

if (-not $result.Found) {
    Write-Output "Policy lookup failed: $PolicyName ($($result.HResultHex))"
    exit 1
}

$items = Split-AllowListValue -Value ([string]$result.Value)

if ($items | Where-Object { $_ -ieq $CheckValue } | Select-Object -First 1) {
    Write-Output "$CheckValue exists in $PolicyName"
    exit 0
}

Write-Output "$CheckValue missing from $PolicyName"
exit 1
