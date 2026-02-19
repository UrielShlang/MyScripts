# Download + Run TeamViewer QuickSupport
$ErrorActionPreference = "Stop"

$url  = "https://150.co.il/downloads/TeamViewerQS14.exe"
$dir  = Join-Path $env:TEMP "TeamViewerQS"
$exe  = Join-Path $dir "TeamViewerQS14.exe"

New-Item -ItemType Directory -Path $dir -Force | Out-Null

Invoke-WebRequest -Uri $url -OutFile $exe -UseBasicParsing

# Run (no UAC prompt expected for QS)
Start-Process -FilePath $exe
