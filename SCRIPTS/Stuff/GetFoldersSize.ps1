$zip = "$env:TEMP\DU.zip"
$dst = "$env:TEMP\DU"
$rawCsv = "$env:TEMP\du_raw.csv"
$filteredCsv = "$env:TEMP\du_over_100mb.csv"

Invoke-WebRequest -Uri "https://download.sysinternals.com/files/DU.zip" -OutFile $zip
Expand-Archive -Path $zip -DestinationPath $dst -Force

& "$dst\du.exe" -nobanner -c -v -q C:\ | Set-Content -Path $rawCsv -Encoding ASCII

$data = Import-Csv -Path $rawCsv | ForEach-Object {
    [PSCustomObject]@{
        Path                = $_.Path
        CurrentFileCount    = [int64]$_.CurrentFileCount
        CurrentFileSize     = [int64]$_.CurrentFileSize
        FileCount           = [int64]$_.FileCount
        DirectoryCount      = [int64]$_.DirectoryCount
        DirectorySize       = [int64]$_.DirectorySize
        DirectorySizeOnDisk = [int64]$_.DirectorySizeOnDisk
        SizeMB              = [math]::Round(([int64]$_.DirectorySize / 1MB), 2)
    }
} | Where-Object { $_.DirectorySize -ge 100MB } |
    Sort-Object DirectorySize -Descending

$data | Export-Csv -Path $filteredCsv -NoTypeInformation -Encoding UTF8
$data | Select-Object Path, SizeMB, DirectorySize | Format-Table -AutoSize

Start-Process $filteredCsv