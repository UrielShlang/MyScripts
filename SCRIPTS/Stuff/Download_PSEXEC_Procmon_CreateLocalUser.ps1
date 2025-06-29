# סקריפט להורדת כלי Sysinternals ויצירת משתמש מקומי
# יש להריץ את הסקריפט כמנהל מערכת

# בדיקה אם הסקריפט רץ כמנהל
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "הסקריפט חייב לרוץ כמנהל מערכת!" -ForegroundColor Red
    Write-Host "לחץ על כל מקש כדי לצאת..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

Write-Host "מתחיל את תהליך ההורדה והתקנה..." -ForegroundColor Green

# חלק 1: הורדת כלי Sysinternals
Write-Host "`n1. מוריד כלי Sysinternals..." -ForegroundColor Yellow

try {
    # יצירת תיקיית Tools אם לא קיימת
    $toolsPath = "C:\Tools"
    if (!(Test-Path $toolsPath)) {
        New-Item -ItemType Directory -Path $toolsPath -Force | Out-Null
        Write-Host "נוצרה תיקייה: $toolsPath" -ForegroundColor Green
    }

    # רשימת הכלים להורדה
    $tools = @(
        @{
            Name = "ProcessMonitor"
            Url = "https://download.sysinternals.com/files/ProcessMonitor.zip"
            Description = "Process Monitor (ProcMon)"
        },
        @{
            Name = "PSTools"
            Url = "https://download.sysinternals.com/files/PSTools.zip"
            Description = "PSExec ועוד כלי PS"
        },
        @{
            Name = "ProcessExplorer"
            Url = "https://download.sysinternals.com/files/ProcessExplorer.zip"
            Description = "Process Explorer"
        },
        @{
            Name = "Autoruns"
            Url = "https://download.sysinternals.com/files/Autoruns.zip"
            Description = "Autoruns"
        }
    )

    foreach ($tool in $tools) {
        Write-Host "מוריד $($tool.Description)..." -ForegroundColor Cyan
        
        $zipPath = "$toolsPath\$($tool.Name).zip"
        $folderPath = "$toolsPath\$($tool.Name)"
        
        # הורדה
        Invoke-WebRequest -Uri $tool.Url -OutFile $zipPath -UseBasicParsing
        
        # חילוץ
        if (Test-Path $folderPath) {
            Remove-Item $folderPath -Recurse -Force
        }
        Expand-Archive -Path $zipPath -DestinationPath $folderPath -Force
        
        # מחיקת קובץ ZIP
        Remove-Item $zipPath -Force
        
        Write-Host "$($tool.Description) הותקן ב: $folderPath" -ForegroundColor Green
    }
    
    Write-Host "`nכל הכלים הותקנו בהצלחה!" -ForegroundColor Green
}
catch {
    Write-Host "שגיאה בהורדת הכלים: $($_.Exception.Message)" -ForegroundColor Red
}

# חלק 2: יצירת משתמש מקומי
Write-Host "`n2. יוצר משתמש מקומי..." -ForegroundColor Yellow

try {
    $userName = "user"
    $password = "123123" # סיסמה ברירת מחדל - מומלץ לשנות
    
    # בדיקה אם המשתמש כבר קיים
    $existingUser = Get-LocalUser -Name $userName -ErrorAction SilentlyContinue
    if ($existingUser) {
        Write-Host "המשתמש '$userName' כבר קיים במערכת" -ForegroundColor Yellow
    }
    else {
        # יצירת המשתמש
        $securePassword = ConvertTo-SecureString $password -AsPlainText -Force
        New-LocalUser -Name $userName -Password $securePassword -Description "משתמש רגיל שנוצר על ידי סקריפט" -PasswordNeverExpires
        
        # הוספה לקבוצת המשתמשים (לא מנהלים)
        Add-LocalGroupMember -Group "Users" -Member $userName
        
        Write-Host "המשתמש '$userName' נוצר בהצלחה!" -ForegroundColor Green
        Write-Host "סיסמה: $password" -ForegroundColor Cyan
        Write-Host "שים לב: מומלץ מאוד לשנות את הסיסמה!" -ForegroundColor Red
    }
}
catch {
    Write-Host "שגיאה ביצירת המשתמש: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nהסקריפט הסתיים!" -ForegroundColor Green
Write-Host "`nמיקומי הכלים:" -ForegroundColor Cyan
Write-Host "• Process Monitor: C:\Tools\ProcessMonitor\Procmon.exe" -ForegroundColor White
Write-Host "• PSExec: C:\Tools\PSTools\PsExec.exe" -ForegroundColor White
Write-Host "• Process Explorer: C:\Tools\ProcessExplorer\procexp.exe" -ForegroundColor White
Write-Host "• Autoruns: C:\Tools\Autoruns\Autoruns.exe" -ForegroundColor White
Write-Host "`nלחץ על כל מקש כדי לצאת..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
