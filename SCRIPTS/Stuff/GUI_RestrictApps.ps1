# GUI Manager for RestrictRun Registry Values
# מנהל GUI עבור ערכי רישום RestrictRun

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# נתיבי רישום אפשריים
$PossibleRegistryPaths = @(
    "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictRun",
    "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictRun",
    "HKEY_USERS\S-1-5-21-2157826234-2034544563-3753416084-1005\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictRun"
)

$RegistryPath = $null

# בדיקת הרשאות אדמין
function Test-AdminRights {
    return ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

# פונקציה לאיתור נתיב הרישום הנכון
function Find-RegistryPath {
    foreach ($path in $PossibleRegistryPaths) {
        try {
            $testPath = "Registry::$path"
            if (Test-Path $testPath) {
                return $path
            }
        }
        catch {
            continue
        }
    }
    return $null
}

# פונקציה ליצירת נתיב הרישום
function Create-RegistryPath {
    param([string]$Path)
    
    try {
        $fullPath = "Registry::$Path"
        if (-not (Test-Path $fullPath)) {
            # יצירת הנתיב המלא
            $pathParts = $Path -split '\\'
            $currentPath = ""
            
            for ($i = 0; $i -lt $pathParts.Length; $i++) {
                if ($i -eq 0) {
                    $currentPath = $pathParts[$i]
                } else {
                    $currentPath = "$currentPath\$($pathParts[$i])"
                }
                
                $testPath = "Registry::$currentPath"
                if (-not (Test-Path $testPath)) {
                    New-Item -Path $testPath -Force | Out-Null
                }
            }
            return $true
        }
        return $true
    }
    catch {
        return $false
    }
}

# אתחול נתיב הרישום
function Initialize-RegistryPath {
    # חיפוש נתיב קיים
    $script:RegistryPath = Find-RegistryPath
    
    if ($script:RegistryPath) {
        return $true
    }
    
    # אם לא נמצא, נסה ליצור ב-HKEY_CURRENT_USER
    $defaultPath = "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\RestrictRun"
    
    if (Create-RegistryPath -Path $defaultPath) {
        $script:RegistryPath = $defaultPath
        return $true
    }
    
    return $false
}

# פונקציות רישום
function Get-RestrictRunValues {
    if (-not $script:RegistryPath) {
        return @()
    }
    
    try {
        $values = Get-ItemProperty -Path "Registry::$script:RegistryPath" -ErrorAction Stop
        $assignedAccess = @()
        
        foreach ($property in $values.PSObject.Properties) {
            if ($property.Name -like "AssignedAccess_*") {
                $number = $property.Name -replace "AssignedAccess_", ""
                $assignedAccess += [PSCustomObject]@{
                    Name = $property.Name
                    Number = [int]$number
                    Value = $property.Value
                }
            }
        }
        
        return $assignedAccess | Sort-Object Number
    }
    catch {
        return @()
    }
}

function Get-NextAssignedAccessNumber {
    $existingValues = Get-RestrictRunValues
    if ($existingValues.Count -eq 0) {
        return 1
    }
    $maxNumber = ($existingValues | Measure-Object -Property Number -Maximum).Maximum
    return $maxNumber + 1
}

function Add-RestrictRunValue {
    param([string]$ExecutableName, [int]$Number = -1)
    
    if (-not $script:RegistryPath) {
        return @{Success = $false; Message = "נתיב הרישום לא זמין"}
    }
    
    try {
        if ($Number -eq -1) {
            $Number = Get-NextAssignedAccessNumber
        }
        
        $valueName = "AssignedAccess_$Number"
        
        # בדיקה אם הערך כבר קיים
        try {
            $existingValue = Get-ItemProperty -Path "Registry::$script:RegistryPath" -Name $valueName -ErrorAction Stop
            return @{Success = $false; Message = "הערך $valueName כבר קיים"}
        }
        catch {
            # זה טוב - הערך לא קיים
        }
        
        Set-ItemProperty -Path "Registry::$script:RegistryPath" -Name $valueName -Value $ExecutableName -Type String
        return @{Success = $true; Message = "נוסף בהצלחה: $valueName = $ExecutableName"}
    }
    catch {
        return @{Success = $false; Message = "שגיאה: $($_.Exception.Message)"}
    }
}

function Remove-RestrictRunValue {
    param([int]$Number)
    
    if (-not $script:RegistryPath) {
        return @{Success = $false; Message = "נתיב הרישום לא זמין"}
    }
    
    $valueName = "AssignedAccess_$Number"
    
    try {
        $existingValue = Get-ItemProperty -Path "Registry::$script:RegistryPath" -Name $valueName -ErrorAction Stop
        Remove-ItemProperty -Path "Registry::$script:RegistryPath" -Name $valueName
        return @{Success = $true; Message = "נמחק בהצלחה: $valueName"}
    }
    catch {
        return @{Success = $false; Message = "שגיאה: הערך לא נמצא או שגיאה במחיקה"}
    }
}

# יצירת GUI
function Create-RestrictRunGUI {
    # בדיקת הרשאות
    if (-not (Test-AdminRights)) {
        [System.Windows.Forms.MessageBox]::Show(
            "נדרשות הרשאות אדמיניסטרטור לעריכת הרישום!`nאנא הפעל את PowerShell כאדמיניסטרטור",
            "שגיאת הרשאות",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # אתחול נתיב הרישום
    if (-not (Initialize-RegistryPath)) {
        [System.Windows.Forms.MessageBox]::Show(
            "שגיאה ביצירת או איתור נתיב הרישום!`nאנא בדוק הרשאות או נסה להריץ כאדמיניסטרטור",
            "שגיאת רישום",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    # חלון ראשי
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "מנהל RestrictRun Registry"
    $form.Size = New-Object System.Drawing.Size(800, 600)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # כותרת
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "מנהל ערכי RestrictRun Registry"
    $titleLabel.Location = New-Object System.Drawing.Point(10, 10)
    $titleLabel.Size = New-Object System.Drawing.Size(760, 30)
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = "MiddleCenter"
    $titleLabel.BackColor = [System.Drawing.Color]::LightBlue
    $form.Controls.Add($titleLabel)

    # קבוצת הוספת ערכים
    $addGroupBox = New-Object System.Windows.Forms.GroupBox
    $addGroupBox.Text = "הוספת ערכים חדשים"
    $addGroupBox.Location = New-Object System.Drawing.Point(10, 50)
    $addGroupBox.Size = New-Object System.Drawing.Size(760, 200)
    $form.Controls.Add($addGroupBox)

    # תיבת טקסט לערכים מרובים
    $valuesLabel = New-Object System.Windows.Forms.Label
    $valuesLabel.Text = "הכנס שמות קבצי הרצה (קובץ אחד בכל שורה):"
    $valuesLabel.Location = New-Object System.Drawing.Point(10, 25)
    $valuesLabel.Size = New-Object System.Drawing.Size(400, 20)
    $addGroupBox.Controls.Add($valuesLabel)

    $valuesTextBox = New-Object System.Windows.Forms.TextBox
    $valuesTextBox.Location = New-Object System.Drawing.Point(10, 50)
    $valuesTextBox.Size = New-Object System.Drawing.Size(400, 100)
    $valuesTextBox.Multiline = $true
    $valuesTextBox.ScrollBars = "Vertical"
    $valuesTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $addGroupBox.Controls.Add($valuesTextBox)

    # תיבת מספר התחלתי
    $startNumberLabel = New-Object System.Windows.Forms.Label
    $startNumberLabel.Text = "מספר התחלתי (אופציונלי):"
    $startNumberLabel.Location = New-Object System.Drawing.Point(430, 25)
    $startNumberLabel.Size = New-Object System.Drawing.Size(150, 20)
    $addGroupBox.Controls.Add($startNumberLabel)

    $startNumberTextBox = New-Object System.Windows.Forms.TextBox
    $startNumberTextBox.Location = New-Object System.Drawing.Point(430, 50)
    $startNumberTextBox.Size = New-Object System.Drawing.Size(100, 25)
    $addGroupBox.Controls.Add($startNumberTextBox)

    # כפתור הוספה
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = "הוסף ערכים"
    $addButton.Location = New-Object System.Drawing.Point(430, 85)
    $addButton.Size = New-Object System.Drawing.Size(100, 30)
    $addButton.BackColor = [System.Drawing.Color]::LightGreen
    $addGroupBox.Controls.Add($addButton)

    # כפתור טקסט לדוגמה
    $exampleButton = New-Object System.Windows.Forms.Button
    $exampleButton.Text = "טען דוגמה"
    $exampleButton.Location = New-Object System.Drawing.Point(430, 125)
    $exampleButton.Size = New-Object System.Drawing.Size(100, 30)
    $exampleButton.BackColor = [System.Drawing.Color]::LightYellow
    $addGroupBox.Controls.Add($exampleButton)

    # כפתור ניקוי
    $clearButton = New-Object System.Windows.Forms.Button
    $clearButton.Text = "נקה טקסט"
    $clearButton.Location = New-Object System.Drawing.Point(540, 85)
    $clearButton.Size = New-Object System.Drawing.Size(100, 30)
    $clearButton.BackColor = [System.Drawing.Color]::LightCoral
    $addGroupBox.Controls.Add($clearButton)

    # כפתור טעינה מ-CSV
    $loadCsvButton = New-Object System.Windows.Forms.Button
    $loadCsvButton.Text = "טען מ-CSV"
    $loadCsvButton.Location = New-Object System.Drawing.Point(540, 125)
    $loadCsvButton.Size = New-Object System.Drawing.Size(100, 30)
    $loadCsvButton.BackColor = [System.Drawing.Color]::LightSalmon
    $addGroupBox.Controls.Add($loadCsvButton)

    # תווית עמודת CSV
    $csvColumnLabel = New-Object System.Windows.Forms.Label
    $csvColumnLabel.Text = "עמודה לטעינה:"
    $csvColumnLabel.Location = New-Object System.Drawing.Point(430, 160)
    $csvColumnLabel.Size = New-Object System.Drawing.Size(100, 20)
    $addGroupBox.Controls.Add($csvColumnLabel)

    # תיבת עמודת CSV
    $csvColumnTextBox = New-Object System.Windows.Forms.TextBox
    $csvColumnTextBox.Location = New-Object System.Drawing.Point(540, 160)
    $csvColumnTextBox.Size = New-Object System.Drawing.Size(100, 25)
    $csvColumnTextBox.Text = "1"
    $addGroupBox.Controls.Add($csvColumnTextBox)

    # הצגת ערכים קיימים
    $existingGroupBox = New-Object System.Windows.Forms.GroupBox
    $existingGroupBox.Text = "ערכים קיימים"
    $existingGroupBox.Location = New-Object System.Drawing.Point(10, 260)
    $existingGroupBox.Size = New-Object System.Drawing.Size(760, 250)
    $form.Controls.Add($existingGroupBox)

    # ListView לערכים קיימים
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(10, 25)
    $listView.Size = New-Object System.Drawing.Size(600, 180)
    $listView.View = "Details"
    $listView.FullRowSelect = $true
    $listView.GridLines = $true
    $listView.Columns.Add("מספר", 80)
    $listView.Columns.Add("שם", 200)
    $listView.Columns.Add("ערך", 300)
    $existingGroupBox.Controls.Add($listView)

    # כפתורי פעולות
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Text = "רענן רשימה"
    $refreshButton.Location = New-Object System.Drawing.Point(620, 25)
    $refreshButton.Size = New-Object System.Drawing.Size(120, 30)
    $refreshButton.BackColor = [System.Drawing.Color]::LightBlue
    $existingGroupBox.Controls.Add($refreshButton)

    $deleteButton = New-Object System.Windows.Forms.Button
    $deleteButton.Text = "מחק נבחר"
    $deleteButton.Location = New-Object System.Drawing.Point(620, 65)
    $deleteButton.Size = New-Object System.Drawing.Size(120, 30)
    $deleteButton.BackColor = [System.Drawing.Color]::LightCoral
    $existingGroupBox.Controls.Add($deleteButton)

    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Text = "יצא לקובץ"
    $exportButton.Location = New-Object System.Drawing.Point(620, 105)
    $exportButton.Size = New-Object System.Drawing.Size(120, 30)
    $exportButton.BackColor = [System.Drawing.Color]::LightGray
    $existingGroupBox.Controls.Add($exportButton)

    # תיבת הודעות
    $statusTextBox = New-Object System.Windows.Forms.TextBox
    $statusTextBox.Location = New-Object System.Drawing.Point(10, 520)
    $statusTextBox.Size = New-Object System.Drawing.Size(760, 40)
    $statusTextBox.Multiline = $true
    $statusTextBox.ReadOnly = $true
    $statusTextBox.ScrollBars = "Vertical"
    $statusTextBox.BackColor = [System.Drawing.Color]::Black
    $statusTextBox.ForeColor = [System.Drawing.Color]::White
    $statusTextBox.Font = New-Object System.Drawing.Font("Consolas", 8)
    $form.Controls.Add($statusTextBox)

    # פונקציות עזר
    function Update-Status {
        param([string]$Message)
        $timestamp = Get-Date -Format "HH:mm:ss"
        $statusTextBox.AppendText("[$timestamp] $Message`r`n")
        $statusTextBox.SelectionStart = $statusTextBox.Text.Length
        $statusTextBox.ScrollToCaret()
        $form.Refresh()
    }

    function Refresh-ListView {
        $listView.Items.Clear()
        $values = Get-RestrictRunValues
        
        foreach ($value in $values) {
            $item = New-Object System.Windows.Forms.ListViewItem($value.Number.ToString())
            $item.SubItems.Add($value.Name)
            $item.SubItems.Add($value.Value)
            $item.Tag = $value.Number
            $listView.Items.Add($item)
        }
        
        Update-Status "נטענו $($values.Count) ערכים"
    }

    function Load-FromCSV {
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $openFileDialog.Title = "בחר קובץ CSV"
        
        if ($openFileDialog.ShowDialog() -eq "OK") {
            try {
                Update-Status "קורא קובץ CSV: $($openFileDialog.FileName)"
                
                # קריאת הקובץ
                $csvContent = Import-Csv -Path $openFileDialog.FileName -Encoding UTF8
                
                if ($csvContent.Count -eq 0) {
                    Update-Status "שגיאה: הקובץ ריק או לא תקין"
                    [System.Windows.Forms.MessageBox]::Show("הקובץ ריק או לא תקין", "שגיאה", "OK", "Warning")
                    return
                }
                
                # הצגת עמודות זמינות
                $columns = $csvContent[0].PSObject.Properties.Name
                Update-Status "עמודות זמינות: $($columns -join ', ')"
                
                # קביעת עמודה לטעינה
                $columnToLoad = $csvColumnTextBox.Text.Trim()
                $selectedColumn = $null
                
                # בדיקה אם זה מספר עמודה
                if ($columnToLoad -match '^\d+$') {
                    $columnIndex = [int]$columnToLoad - 1
                    if ($columnIndex -ge 0 -and $columnIndex -lt $columns.Count) {
                        $selectedColumn = $columns[$columnIndex]
                    }
                }
                # בדיקה אם זה שם עמודה
                elseif ($columns -contains $columnToLoad) {
                    $selectedColumn = $columnToLoad
                }
                # חיפוש עמודה דומה
                else {
                    $similarColumn = $columns | Where-Object { $_ -like "*$columnToLoad*" } | Select-Object -First 1
                    if ($similarColumn) {
                        $selectedColumn = $similarColumn
                        Update-Status "נמצאה עמודה דומה: $similarColumn"
                    }
                }
                
                if (-not $selectedColumn) {
                    $columnsList = for ($i = 0; $i -lt $columns.Count; $i++) {
                        "$($i + 1). $($columns[$i])"
                    }
                    $message = "לא נמצאה עמודה '$columnToLoad'`n`nעמודות זמינות:`n$($columnsList -join "`n")"
                    [System.Windows.Forms.MessageBox]::Show($message, "שגיאת עמודה", "OK", "Warning")
                    return
                }
                
                # חילוץ הערכים מהעמודה הנבחרת
                $values = @()
                foreach ($row in $csvContent) {
                    $value = $row.$selectedColumn
                    if ($value -and $value.ToString().Trim() -ne "") {
                        # אם זה נתיב מלא, חלץ רק את שם הקובץ
                        if ($value -match '\\') {
                            $value = Split-Path -Leaf $value
                        }
                        
                        # הוסף .exe אם חסר
                        if ($value -notmatch '\.\w+$') {
                            $value = "$value.exe"
                        }
                        
                        $values += $value.ToString().Trim()
                    }
                }
                
                # הסרת כפילויות
                $uniqueValues = $values | Sort-Object -Unique
                
                if ($uniqueValues.Count -eq 0) {
                    Update-Status "שגיאה: לא נמצאו ערכים תקינים בעמודה '$selectedColumn'"
                    [System.Windows.Forms.MessageBox]::Show("לא נמצאו ערכים תקינים בעמודה הנבחרת", "שגיאה", "OK", "Warning")
                    return
                }
                
                # עדכון תיבת הטקסט
                $valuesTextBox.Text = $uniqueValues -join "`r`n"
                
                Update-Status "נטענו $($uniqueValues.Count) ערכים ייחודיים מעמודה '$selectedColumn'"
                Update-Status "ערכים: $($uniqueValues -join ', ')"
                
                [System.Windows.Forms.MessageBox]::Show(
                    "נטענו בהצלחה $($uniqueValues.Count) שמות קבצים מהעמודה '$selectedColumn'",
                    "טעינה הושלמה",
                    "OK",
                    "Information"
                )
            }
            catch {
                $errorMsg = "שגיאה בטעינת קובץ CSV: $($_.Exception.Message)"
                Update-Status $errorMsg
                [System.Windows.Forms.MessageBox]::Show($errorMsg, "שגיאה", "OK", "Error")
            }
        }
    }

    # אירועים
    $addButton.Add_Click({
        $inputText = $valuesTextBox.Text.Trim()
        if ([string]::IsNullOrEmpty($inputText)) {
            [System.Windows.Forms.MessageBox]::Show("אנא הכנס לפחות שם קובץ אחד", "שגיאה", "OK", "Warning")
            return
        }

        $executables = $inputText -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        $startNumber = if ($startNumberTextBox.Text -match '^\d+$') { [int]$startNumberTextBox.Text } else { -1 }
        
        $successCount = 0
        $currentNumber = $startNumber
        
        foreach ($executable in $executables) {
            $result = Add-RestrictRunValue -ExecutableName $executable -Number $currentNumber
            
            if ($result.Success) {
                Update-Status "✓ $($result.Message)"
                $successCount++
            } else {
                Update-Status "✗ $($result.Message)"
            }
            
            if ($currentNumber -ne -1) {
                $currentNumber++
            }
        }
        
        Update-Status "הושלמה הוספת $successCount מתוך $($executables.Count) ערכים"
        Refresh-ListView
    })

    $exampleButton.Add_Click({
        $valuesTextBox.Text = @"
notepad.exe
calc.exe
msedge.exe
chrome.exe
firefox.exe
"@
    })

    $clearButton.Add_Click({
        $valuesTextBox.Clear()
        $startNumberTextBox.Clear()
    })

    $loadCsvButton.Add_Click({
        Load-FromCSV
    })

    $refreshButton.Add_Click({
        Refresh-ListView
    })

    $deleteButton.Add_Click({
        if ($listView.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("אנא בחר ערך למחיקה", "שגיאה", "OK", "Warning")
            return
        }

        $result = [System.Windows.Forms.MessageBox]::Show(
            "האם אתה בטוח שברצונך למחוק את הערכים הנבחרים?",
            "אישור מחיקה",
            "YesNo",
            "Question"
        )

        if ($result -eq "Yes") {
            foreach ($selectedItem in $listView.SelectedItems) {
                $number = $selectedItem.Tag
                $deleteResult = Remove-RestrictRunValue -Number $number
                
                if ($deleteResult.Success) {
                    Update-Status "✓ $($deleteResult.Message)"
                } else {
                    Update-Status "✗ $($deleteResult.Message)"
                }
            }
            Refresh-ListView
        }
    })

    $exportButton.Add_Click({
        $values = Get-RestrictRunValues
        if ($values.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("אין ערכים לייצא", "שגיאה", "OK", "Warning")
            return
        }

        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        $saveDialog.FileName = "RestrictRun_Values_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"

        if ($saveDialog.ShowDialog() -eq "OK") {
            try {
                $exportContent = "RestrictRun Registry Values - $(Get-Date)`n"
                $exportContent += "=" * 50 + "`n"
                
                foreach ($value in $values) {
                    $exportContent += "$($value.Name) = $($value.Value)`n"
                }
                
                [System.IO.File]::WriteAllText($saveDialog.FileName, $exportContent, [System.Text.Encoding]::UTF8)
                Update-Status "נשמר בהצלחה: $($saveDialog.FileName)"
            }
            catch {
                Update-Status "שגיאה בשמירה: $($_.Exception.Message)"
            }
        }
    })

    # טעינה ראשונית
    Update-Status "מנהל RestrictRun הופעל בהצלחה"
    Update-Status "נתיב רישום: $script:RegistryPath"
    Refresh-ListView

    # הצגת החלון
    $form.Add_Shown({$form.Activate()})
    [void]$form.ShowDialog()
}

# הפעלת GUI
Create-RestrictRunGUI
