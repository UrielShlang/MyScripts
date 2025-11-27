<#
.SYNOPSIS
    סורק את כל תיבות הדואר האישיות והמשותפות ב-Exchange Online ומחזיר דוח גודל ואחוז מילוי.

.DESCRIPTION

    הסקריפט מתחבר ל-Exchange Online וסורק:
    - תיבות דואר אישיות (UserMailbox)
    - תיבות דואר משותפות (SharedMailbox)
    
    עבור כל תיבה מוחזר:
    - שם התצוגה
    - כתובת המייל
    - סוג התיבה
    - גודל נוכחי
    - מכסה (Quota)
    - אחוז מילוי

.PARAMETER ExportPath
    נתיב לייצוא הדוח ל-CSV (אופציונלי)

.EXAMPLE
    .\Get-MailboxSizeReport.ps1
    .\Get-MailboxSizeReport.ps1 -ExportPath "C:\Reports\MailboxReport.csv"

.NOTES
    דרישות: מודול ExchangeOnlineManagement
    הרשאות: Exchange Administrator או Global Reader
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ExportPath
)

#region Functions

function Convert-ToBytes {
    param([string]$SizeString)
    
    if ([string]::IsNullOrWhiteSpace($SizeString)) { return 0 }
    
    # מחלץ את הגודל מפורמט כמו "1.234 GB (1,234,567,890 bytes)"
    if ($SizeString -match '\(([0-9,]+)\s*bytes\)') {
        return [long]($Matches[1] -replace ',', '')
    }
    
    # פורמט פשוט עם יחידות
    $SizeString = $SizeString.Trim()
    
    switch -Regex ($SizeString) {
        '([0-9.,]+)\s*TB' { return [math]::Round([double]($Matches[1] -replace ',', '.') * 1TB) }
        '([0-9.,]+)\s*GB' { return [math]::Round([double]($Matches[1] -replace ',', '.') * 1GB) }
        '([0-9.,]+)\s*MB' { return [math]::Round([double]($Matches[1] -replace ',', '.') * 1MB) }
        '([0-9.,]+)\s*KB' { return [math]::Round([double]($Matches[1] -replace ',', '.') * 1KB) }
        '([0-9.,]+)\s*B'  { return [long]($Matches[1] -replace ',', '') }
        default { return 0 }
    }
}

function Format-Size {
    param([long]$Bytes)
    
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Get-EffectiveQuota {
    param($Mailbox)
    
    # בודק אם יש הגדרת מכסה ספציפית לתיבה או שמשתמשים בברירת מחדל
    $quota = $Mailbox.ProhibitSendReceiveQuota
    
    if ($quota -eq 'unlimited' -or [string]::IsNullOrWhiteSpace($quota)) {
        # ברירת מחדל לתיבות רגילות: 50GB, לתיבות משותפות: 50GB
        return 50GB
    }
    
    return Convert-ToBytes $quota.ToString()
}

#endregion

#region Main

Write-Host "`n===== דוח גודל תיבות דואר - Exchange Online =====" -ForegroundColor Cyan
Write-Host "התחלת סריקה: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')`n" -ForegroundColor Gray

# בדיקת חיבור ל-Exchange Online
try {
    $null = Get-OrganizationConfig -ErrorAction Stop
    Write-Host "[✓] מחובר ל-Exchange Online" -ForegroundColor Green
}
catch {
    Write-Host "[!] לא מחובר ל-Exchange Online. מתחבר..." -ForegroundColor Yellow
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Host "[✓] התחברות הצליחה" -ForegroundColor Green
    }
    catch {
        Write-Error "שגיאה בהתחברות ל-Exchange Online: $_"
        exit 1
    }
}

# שליפת כל התיבות
Write-Host "`n[*] שולף תיבות דואר..." -ForegroundColor Cyan

$mailboxes = @()

# תיבות אישיות
Write-Host "    - תיבות אישיות..." -NoNewline
$userMailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited
Write-Host " נמצאו $($userMailboxes.Count)" -ForegroundColor Green
$mailboxes += $userMailboxes

# תיבות משותפות
Write-Host "    - תיבות משותפות..." -NoNewline
$sharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
Write-Host " נמצאו $($sharedMailboxes.Count)" -ForegroundColor Green
$mailboxes += $sharedMailboxes

Write-Host "`n[*] סורק סטטיסטיקות עבור $($mailboxes.Count) תיבות..." -ForegroundColor Cyan

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$counter = 0

foreach ($mbx in $mailboxes) {
    $counter++
    $percent = [math]::Round(($counter / $mailboxes.Count) * 100)
    Write-Progress -Activity "סורק תיבות דואר" -Status "$counter מתוך $($mailboxes.Count) - $($mbx.DisplayName)" -PercentComplete $percent
    
    try {
        $stats = Get-MailboxStatistics -Identity $mbx.Identity -ErrorAction Stop
        $currentSizeBytes = Convert-ToBytes $stats.TotalItemSize.ToString()
        $quotaBytes = Get-EffectiveQuota $mbx
        
        $percentUsed = if ($quotaBytes -gt 0) {
            [math]::Round(($currentSizeBytes / $quotaBytes) * 100, 2)
        } else { 0 }
        
        $results.Add([PSCustomObject]@{
            DisplayName     = $mbx.DisplayName
            EmailAddress    = $mbx.PrimarySmtpAddress
            MailboxType     = if ($mbx.RecipientTypeDetails -eq 'UserMailbox') { 'אישית' } else { 'משותפת' }
            CurrentSize     = Format-Size $currentSizeBytes
            CurrentSizeBytes = $currentSizeBytes
            Quota           = Format-Size $quotaBytes
            QuotaBytes      = $quotaBytes
            PercentUsed     = $percentUsed
            ItemCount       = $stats.ItemCount
            LastLogonTime   = $stats.LastLogonTime
        })
    }
    catch {
        Write-Warning "שגיאה בקריאת סטטיסטיקות עבור $($mbx.DisplayName): $_"
        $results.Add([PSCustomObject]@{
            DisplayName     = $mbx.DisplayName
            EmailAddress    = $mbx.PrimarySmtpAddress
            MailboxType     = if ($mbx.RecipientTypeDetails -eq 'UserMailbox') { 'אישית' } else { 'משותפת' }
            CurrentSize     = 'שגיאה'
            CurrentSizeBytes = 0
            Quota           = 'N/A'
            QuotaBytes      = 0
            PercentUsed     = 0
            ItemCount       = 0
            LastLogonTime   = $null
        })
    }
}

Write-Progress -Activity "סורק תיבות דואר" -Completed

# מיון לפי אחוז מילוי (מהגבוה לנמוך)
$sortedResults = $results | Sort-Object -Property PercentUsed -Descending

# הצגת תוצאות
Write-Host "`n===== תוצאות =====" -ForegroundColor Cyan

# סיכום
$totalSize = ($results | Measure-Object -Property CurrentSizeBytes -Sum).Sum
$userCount = ($results | Where-Object { $_.MailboxType -eq 'אישית' }).Count
$sharedCount = ($results | Where-Object { $_.MailboxType -eq 'משותפת' }).Count
$over80 = ($results | Where-Object { $_.PercentUsed -ge 80 }).Count
$over90 = ($results | Where-Object { $_.PercentUsed -ge 90 }).Count

Write-Host "`nסיכום:" -ForegroundColor Yellow
Write-Host "  סה״כ תיבות: $($results.Count) (אישיות: $userCount, משותפות: $sharedCount)"
Write-Host "  נפח כולל: $(Format-Size $totalSize)"
Write-Host "  תיבות מעל 80%: $over80" -ForegroundColor $(if ($over80 -gt 0) { 'Yellow' } else { 'Green' })
Write-Host "  תיבות מעל 90%: $over90" -ForegroundColor $(if ($over90 -gt 0) { 'Red' } else { 'Green' })

# טבלה
Write-Host "`nרשימת תיבות (ממוינות לפי אחוז מילוי):" -ForegroundColor Yellow

$sortedResults | Select-Object DisplayName, EmailAddress, MailboxType, CurrentSize, Quota, 
    @{N='PercentUsed';E={"$($_.PercentUsed)%"}}, ItemCount | 
    Format-Table -AutoSize

# ייצוא ל-CSV
if ($ExportPath) {
    try {
        $sortedResults | Select-Object DisplayName, EmailAddress, MailboxType, CurrentSize, 
            CurrentSizeBytes, Quota, QuotaBytes, PercentUsed, ItemCount, LastLogonTime |
            Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
        Write-Host "`n[✓] הדוח יוצא ל: $ExportPath" -ForegroundColor Green
    }
    catch {
        Write-Error "שגיאה בייצוא הדוח: $_"
    }
}

# החזרת התוצאות
Write-Host "`nסיום סריקה: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor Gray
return $sortedResults

#endregion
#$sortedResults | Out-GridView
