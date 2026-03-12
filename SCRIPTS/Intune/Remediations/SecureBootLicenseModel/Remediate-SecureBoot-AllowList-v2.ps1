#Requires -RunAsAdministrator
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-CommandPath {
    param(
        [Parameter(Mandatory)]
        [string]$FileName
    )

    $candidates = @(
        (Join-Path $env:SystemRoot "System32\$FileName"),
        (Join-Path $env:SystemRoot "Sysnative\$FileName"),
        $FileName
    )

    foreach ($candidate in $candidates) {
        try {
            if ($candidate -eq $FileName) {
                $cmd = Get-Command $FileName -ErrorAction Stop
                return $cmd.Source
            }

            if (Test-Path -LiteralPath $candidate) {
                return $candidate
            }
        }
        catch {
        }
    }

    throw "Could not find $FileName"
}

$clipDls = Resolve-CommandPath -FileName 'ClipDLS.exe'
$clipRenew = Resolve-CommandPath -FileName 'ClipRenew.exe'

& $clipDls removesubscription
if ($LASTEXITCODE -ne 0) {
    throw "ClipDLS.exe failed with exit code $LASTEXITCODE"
}

Start-Sleep -Seconds 30

& $clipRenew
if ($LASTEXITCODE -ne 0) {
    throw "ClipRenew.exe failed with exit code $LASTEXITCODE"
}
Start-Sleep -Seconds 30

exit 0
