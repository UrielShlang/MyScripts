Add-Type -AssemblyName System.Windows.Forms

# Conversion Function (unchanged)
function Convert-Indicators($inputFile, $outputFile) {
    $data = Import-Csv -Path $inputFile
    $results = @()

    foreach ($row in $data) {
        foreach ($col in @('sha1','sha256','IP','domain','url')) {
            if ($row.$col) {
                $indicatorType = switch ($col) {
                    'sha1'    { 'FileSha1' }
                    'sha256'  { 'FileSha256' }
                    'IP'      { 'IpAddress' }
                    'domain'  { 'DomainName' }
                    'url'     { 'Url' }
                }

                $action = if ($col -eq 'IP') { 'Block' }
                          elseif ($col -in 'sha1','sha256') { 'BlockAndRemediate' }
                          else { 'Block' }
                $value = $row.$col.Replace('[.]','.').Replace('hxxp','http').Replace('HXXP','HTTP')


                $description = "Report by CERT " + $row.publishDate

                $results += [PSCustomObject]@{
                    IndicatorType      = $indicatorType
                    IndicatorValue     = $value
                    ExpirationTime     = ''
                    Action             = $action
                    Severity           = 'High'
                    Title              = if ($row.title) { $row.title } else { 'Imported Indicator' }
                    Description        = $description
                    RecommendedActions = 'Review and take necessary actions'
                    RbacGroups         = ''
                    Category           = 'SuspiciousActivity'
                    MitreTechniques    = ''
                    GenerateAlert      = 'True'
                }
            }
        }
    }

    $results | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8
}

# UI Elements (unchanged)
$form = New-Object System.Windows.Forms.Form
$form.Text = "TI Indicator Converter"
$form.Width = 400
$form.Height = 200
$form.StartPosition = "CenterScreen"

$btnUpload = New-Object System.Windows.Forms.Button
$btnUpload.Location = New-Object System.Drawing.Point(120, 40)
$btnUpload.Size = New-Object System.Drawing.Size(150, 30)
$btnUpload.Text = "Upload CSV"
$form.Controls.Add($btnUpload)

$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "CSV files (*.csv)|*.csv"

$saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$saveFileDialog.Filter = "CSV files (*.csv)|*.csv"

$btnUpload.Add_Click({
    if ($openFileDialog.ShowDialog() -eq "OK" -and $saveFileDialog.ShowDialog() -eq "OK") {
        Convert-Indicators -inputFile $openFileDialog.FileName `
                           -outputFile $saveFileDialog.FileName
        [System.Windows.Forms.MessageBox]::Show("File Converted Successfully!", "Success")
    }
})

$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
#[void]$form.ShowDialog()

# ----------------------------------------------------------------
# Automatic folder processing
# ----------------------------------------------------------------

$folderPath = 'C:\Users\UrielShlang\OneDrive - Pro-Vision Information Systems Ltd\מסמכים\CERT'

Get-ChildItem -Path $folderPath -Filter 'ALERT-CERT-IL-W-*.csv' |
    Where-Object { $_.Name -notmatch '_Converted\.csv$' } |
    ForEach-Object {
        $inputFile  = $_.FullName
        $baseName   = $_.BaseName
        $ext        = $_.Extension
        $outputFile = Join-Path $_.DirectoryName ("$baseName`_Converted$ext")

        if (-not (Test-Path $outputFile)) {
            Convert-Indicators -inputFile $inputFile -outputFile $outputFile
            Write-Host "Converted: $($_.Name) -> $([System.IO.Path]::GetFileName($outputFile))"
        }
        else {
            Write-Host "Skipping (already exists): $([System.IO.Path]::GetFileName($outputFile))"
        }
    }
