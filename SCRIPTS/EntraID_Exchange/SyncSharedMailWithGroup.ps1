# Exchange Management GUI - Compact Version
# גרסה מקוצרת עם כל התכונות העיקריות

#Requires -Version 5.1
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# משתנים גלובליים
$Global:ExchangeConnected = $false
$Global:SharedMailboxes = @()
$Global:AllGroups = @()



# Connect to Microsoft Graph

# Global Graph Connection status
$Global:GraphConnected = $false
function Connect-ToGraph {
    try {
        Update-Status "מתחבר ל־Graph..." "Blue"
        Connect-MgGraph -Scopes Group.Read.All,User.Read.All
        if (Get-MgContext) {
            $Global:GraphConnected = $true
            Update-Status "מחובר ל־Microsoft Graph" "Green"
            $graphButton.Text = "מחובר ל־Graph"
            $graphButton.Enabled = $false
            $graphButton.BackColor = "LightGreen"
        }
    }
    catch {
        Update-Status "שגיאה בחיבור ל־Graph: $($_.Exception.Message)" "Red"
        [System.Windows.Forms.MessageBox]::Show("שגיאה בחיבור ל־Graph: $($_.Exception.Message)", "שגיאה")
    }
}



# פונקציות עזר
function Test-ExchangeConnection {
    try { Get-Mailbox -ResultSize 1 -ErrorAction Stop | Out-Null; return $true } catch { return $false }
}

function Update-Status($Message, $Color = "Black") {
    $statusLabel.Text = $Message
    $statusLabel.ForeColor = [System.Drawing.Color]::FromName($Color)
    $statusLabel.Refresh()
}

# התחברות ל-Exchange
function Connect-ToExchange {
    try {
        $form.Enabled = $false
        Update-Status "מתחבר ל-Exchange..." "Blue"
        
        if (Test-ExchangeConnection) {
            Enable-ExchangeFeatures
            Update-Status "מחובר ל-Exchange (חיבור קיים)" "Green"
            Load-SharedMailboxes
            return
        }
        
        if (-not (Get-Module -ListAvailable ExchangeOnlineManagement)) {
            Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
        }
        Import-Module ExchangeOnlineManagement -Force
        Connect-ExchangeOnline -ShowProgress $false -ShowBanner:$false
        
        if (Test-ExchangeConnection) {
            Enable-ExchangeFeatures
            Update-Status "מחובר ל-Exchange Online בהצלחה" "Green"
            Load-SharedMailboxes
        }
    }
    catch {
        Update-Status "שגיאה בהתחברות: $($_.Exception.Message)" "Red"
        [System.Windows.Forms.MessageBox]::Show("שגיאה: $($_.Exception.Message)", "שגיאה")
    }
    finally { $form.Enabled = $true }
}

function Enable-ExchangeFeatures {
    $connectButton.Text = "מחובר"; $connectButton.Enabled = $false; $connectButton.BackColor = "LightGreen"
    $refreshButton.Enabled = $true; $createMailboxButton.Enabled = $true
    $createGroupButton.Enabled = $true; $syncPermissionsButton.Enabled = $true
}

function Disconnect-Exchange {
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    $connectButton.Text = "התחבר ל-Exchange"; $connectButton.Enabled = $true; $connectButton.BackColor = "LightBlue"
    $refreshButton.Enabled = $false; $createMailboxButton.Enabled = $false
    $createGroupButton.Enabled = $false; $syncPermissionsButton.Enabled = $false
    $mailboxListBox.Items.Clear(); Update-Status "מנותק מ-Exchange" "Orange"
}

# זיהוי תיבות של משתמשים שעזבו
function Test-IsLeaverMailbox($Mailbox) {
    $patterns = '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', '^[0-9a-f-]{30,}$', '^[a-f0-9]{20,}$'
    foreach ($pattern in $patterns) {
        if ($Mailbox.Name -match $pattern -or $Mailbox.Alias -match $pattern) { return $true }
    }
    return ([System.Guid]::TryParse($Mailbox.Name, [ref][System.Guid]::Empty) -or $Mailbox.Name.Length -gt 15 -and $Mailbox.Name -match '^[a-f0-9]+$')
}

# טעינת תיבות דואר משותפות
function Load-SharedMailboxes {
    try {
        Update-Status "טוען תיבות דואר משותפות..." "Blue"
        $allMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited
        $Global:SharedMailboxes = $allMailboxes | Where-Object { -not (Test-IsLeaverMailbox $_) }
        
        $mailboxListBox.Items.Clear()
        foreach ($mb in ($Global:SharedMailboxes | Sort-Object DisplayName)) {
            $mailboxListBox.Items.Add("📧 $($mb.DisplayName) - $($mb.PrimarySmtpAddress)")
        }
        Update-Status "נטענו $($Global:SharedMailboxes.Count) תיבות דואר משותפות" "Green"
    }
    catch { Update-Status "שגיאה בטעינת תיבות דואר: $($_.Exception.Message)" "Red" }
}

# טעינת קבוצות (כולל Entra ID Security Groups)
# Load Groups via Graph API
function Load-AllGroups {
    if (-not $Global:GraphConnected) {
        [System.Windows.Forms.MessageBox]::Show("התחבר קודם ל־Microsoft Graph.", "שגיאה")
        return
    }

    try {
        Update-Status "טוען קבוצות מ־Graph..." "Blue"

        $Global:AllGroups = @()

        # Microsoft 365 Groups
        $m365Groups = Get-MgGroup -Filter "groupTypes/any(c:c eq 'Unified')" -All | Select-Object @{N="Name";E={$_.DisplayName}},@{N="DisplayName";E={$_.DisplayName}},@{N="PrimarySmtpAddress";E={$_.Mail}},@{N="Identity";E={$_.Id}},@{N="GroupType";E={"Microsoft 365 Group"}},@{N="Source";E={"Graph"}}

        # Security Groups
        $securityGroups = Get-MgGroup -Filter "securityEnabled eq true" -All | Select-Object @{N="Name";E={$_.DisplayName}},@{N="DisplayName";E={$_.DisplayName}},@{N="PrimarySmtpAddress";E={$_.Mail}},@{N="Identity";E={$_.Id}},@{N="GroupType";E={"Security Group"}},@{N="Source";E={"Graph"}}

        $Global:AllGroups += $m365Groups + $securityGroups

        Write-Host "✅ נטענו $($m365Groups.Count) Microsoft 365 Groups ו־$($securityGroups.Count) Security Groups" -ForegroundColor Green

        return $Global:AllGroups
    }
    catch {
        Write-Host "❌ שגיאה בטעינת קבוצות מ־Graph: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# פונקציה לקבלת חברי קבוצה לפי מקור
# Get members via Graph
function Get-GroupMembers($GroupObject) {
    try {
        $members = Get-MgGroupMember -GroupId $GroupObject.Identity -All | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.user' }
        $members | Select-Object @{N="DisplayName";E={$_.AdditionalProperties.displayName}}, @{N="PrimarySmtpAddress";E={$_.AdditionalProperties.mail}}, @{N="UserPrincipalName";E={$_.AdditionalProperties.userPrincipalName}}
    }
    catch {
        Write-Host "❌ שגיאה בקבלת חברי קבוצה מ־Graph: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

# יצירת תיבה משותפת
function Create-SharedMailbox {
    $form2 = New-Object System.Windows.Forms.Form
    $form2.Text = "יצירת תיבת דואר משותפת"; $form2.Size = "450,300"; $form2.StartPosition = "CenterParent"
    $form2.RightToLeft = "Yes"; $form2.FormBorderStyle = "FixedDialog"
    
    $nameLabel = New-Object System.Windows.Forms.Label; $nameLabel.Location = "320,20"; $nameLabel.Text = "שם התיבה:"; $form2.Controls.Add($nameLabel)
    $nameBox = New-Object System.Windows.Forms.TextBox; $nameBox.Location = "20,20"; $nameBox.Size = "290,20"; $form2.Controls.Add($nameBox)
    
    $emailLabel = New-Object System.Windows.Forms.Label; $emailLabel.Location = "320,50"; $emailLabel.Text = "אימייל:"; $form2.Controls.Add($emailLabel)
    $emailBox = New-Object System.Windows.Forms.TextBox; $emailBox.Location = "20,50"; $emailBox.Size = "290,20"; $form2.Controls.Add($emailBox)
    
    $displayLabel = New-Object System.Windows.Forms.Label; $displayLabel.Location = "320,80"; $displayLabel.Text = "שם תצוגה:"; $form2.Controls.Add($displayLabel)
    $displayBox = New-Object System.Windows.Forms.TextBox; $displayBox.Location = "20,80"; $displayBox.Size = "290,20"; $form2.Controls.Add($displayBox)
    
    $createBtn = New-Object System.Windows.Forms.Button; $createBtn.Location = "230,200"; $createBtn.Size = "80,30"
    $createBtn.Text = "צור"; $createBtn.BackColor = "LightGreen"; $form2.Controls.Add($createBtn)
    
    $cancelBtn = New-Object System.Windows.Forms.Button; $cancelBtn.Location = "130,200"; $cancelBtn.Size = "80,30"
    $cancelBtn.Text = "בטל"; $cancelBtn.DialogResult = "Cancel"; $form2.Controls.Add($cancelBtn)
    
    $createBtn.Add_Click({
        if ($nameBox.Text -and $emailBox.Text -and $displayBox.Text) {
            try {
                New-Mailbox -Name $nameBox.Text -Alias $emailBox.Text.Split('@')[0] -PrimarySmtpAddress $emailBox.Text -DisplayName $displayBox.Text -Shared
                [System.Windows.Forms.MessageBox]::Show("תיבה נוצרה בהצלחה!", "הצלחה")
                $form2.Close(); Load-SharedMailboxes
            }
            catch { [System.Windows.Forms.MessageBox]::Show("שגיאה: $($_.Exception.Message)", "שגיאה") }
        }
        else { [System.Windows.Forms.MessageBox]::Show("מלא את כל השדות", "שגיאה") }
    })
    
    $form2.ShowDialog(); $form2.Dispose()
}

# יצירת קבוצה
function Create-Group {
    $form2 = New-Object System.Windows.Forms.Form
    $form2.Text = "יצירת קבוצת הפצה"; $form2.Size = "450,350"; $form2.StartPosition = "CenterParent"
    $form2.RightToLeft = "Yes"; $form2.FormBorderStyle = "FixedDialog"
    
    $nameLabel = New-Object System.Windows.Forms.Label; $nameLabel.Location = "320,20"; $nameLabel.Text = "שם הקבוצה:"; $form2.Controls.Add($nameLabel)
    $nameBox = New-Object System.Windows.Forms.TextBox; $nameBox.Location = "20,20"; $nameBox.Size = "290,20"; $form2.Controls.Add($nameBox)
    
    $emailLabel = New-Object System.Windows.Forms.Label; $emailLabel.Location = "320,50"; $emailLabel.Text = "אימייל:"; $form2.Controls.Add($emailLabel)
    $emailBox = New-Object System.Windows.Forms.TextBox; $emailBox.Location = "20,50"; $emailBox.Size = "290,20"; $form2.Controls.Add($emailBox)
    
    $displayLabel = New-Object System.Windows.Forms.Label; $displayLabel.Location = "320,80"; $displayLabel.Text = "שם תצוגה:"; $form2.Controls.Add($displayLabel)
    $displayBox = New-Object System.Windows.Forms.TextBox; $displayBox.Location = "20,80"; $displayBox.Size = "290,20"; $form2.Controls.Add($displayBox)
    
    $typeLabel = New-Object System.Windows.Forms.Label; $typeLabel.Location = "320,110"; $typeLabel.Text = "סוג:"; $form2.Controls.Add($typeLabel)
    $typeBox = New-Object System.Windows.Forms.ComboBox; $typeBox.Location = "20,110"; $typeBox.Size = "290,20"
    $typeBox.Items.AddRange(@("Distribution", "Security")); $typeBox.SelectedIndex = 0; $form2.Controls.Add($typeBox)
    
    $createBtn = New-Object System.Windows.Forms.Button; $createBtn.Location = "230,250"; $createBtn.Size = "80,30"
    $createBtn.Text = "צור"; $createBtn.BackColor = "LightGreen"; $form2.Controls.Add($createBtn)
    
    $cancelBtn = New-Object System.Windows.Forms.Button; $cancelBtn.Location = "130,250"; $cancelBtn.Size = "80,30"
    $cancelBtn.Text = "בטל"; $cancelBtn.DialogResult = "Cancel"; $form2.Controls.Add($cancelBtn)
    
    $createBtn.Add_Click({
        if ($nameBox.Text -and $emailBox.Text -and $displayBox.Text) {
            try {
                New-DistributionGroup -Name $nameBox.Text -Alias $emailBox.Text.Split('@')[0] -PrimarySmtpAddress $emailBox.Text -DisplayName $displayBox.Text -Type $typeBox.SelectedItem
                [System.Windows.Forms.MessageBox]::Show("קבוצה נוצרה בהצלחה!", "הצלחה")
                $form2.Close()
            }
            catch { [System.Windows.Forms.MessageBox]::Show("שגיאה: $($_.Exception.Message)", "שגיאה") }
        }
        else { [System.Windows.Forms.MessageBox]::Show("מלא את כל השדות", "שגיאה") }
    })
    
    $form2.ShowDialog(); $form2.Dispose()
}

# סנכרון הרשאות
function Sync-Permissions {
    $syncForm = New-Object System.Windows.Forms.Form
    $syncForm.Text = "סנכרון הרשאות"; $syncForm.Size = "600,500"; $syncForm.StartPosition = "CenterParent"
    $syncForm.RightToLeft = "Yes"; $syncForm.FormBorderStyle = "FixedDialog"
    
    $groupLabel = New-Object System.Windows.Forms.Label; $groupLabel.Location = "480,20"; $groupLabel.Text = "קבוצה:"; $syncForm.Controls.Add($groupLabel)
    $groupBox = New-Object System.Windows.Forms.ComboBox; $groupBox.Location = "20,20"; $groupBox.Size = "450,20"; $syncForm.Controls.Add($groupBox)
    
    $mailboxLabel = New-Object System.Windows.Forms.Label; $mailboxLabel.Location = "480,50"; $mailboxLabel.Text = "תיבת דואר:"; $syncForm.Controls.Add($mailboxLabel)
    $mailboxBox = New-Object System.Windows.Forms.ComboBox; $mailboxBox.Location = "20,50"; $mailboxBox.Size = "450,20"; $syncForm.Controls.Add($mailboxBox)
    
    $permLabel = New-Object System.Windows.Forms.Label; $permLabel.Location = "480,80"; $permLabel.Text = "הרשאה:"; $syncForm.Controls.Add($permLabel)
    $permBox = New-Object System.Windows.Forms.ComboBox; $permBox.Location = "20,80"; $permBox.Size = "200,20"
    $permBox.Items.AddRange(@("FullAccess", "ReadPermission", "SendAs")); $permBox.SelectedIndex = 0; $syncForm.Controls.Add($permBox)
    
    $resultsBox = New-Object System.Windows.Forms.ListBox; $resultsBox.Location = "20,120"; $resultsBox.Size = "550,250"; $syncForm.Controls.Add($resultsBox)
    
    $syncBtn = New-Object System.Windows.Forms.Button; $syncBtn.Location = "400,390"; $syncBtn.Size = "80,30"
    $syncBtn.Text = "סנכרן"; $syncBtn.BackColor = "LightGreen"; $syncForm.Controls.Add($syncBtn)
    
    $closeBtn = New-Object System.Windows.Forms.Button; $closeBtn.Location = "300,390"; $closeBtn.Size = "80,30"
    $closeBtn.Text = "סגור"; $closeBtn.DialogResult = "Cancel"; $syncForm.Controls.Add($closeBtn)
    
    # טעינת נתונים
    $resultsBox.Items.Add("🔄 טוען קבוצות מכל המקורות...")
    $resultsBox.Refresh()
    
    $groups = Load-AllGroups
    
    $resultsBox.Items.Clear()
    if ($groups -and $groups.Count -gt 0) {
        foreach ($g in ($groups | Sort-Object GroupType, DisplayName)) {
            $groupBox.Items.Add("$($g.DisplayName) [$($g.GroupType)]")
        }
        $resultsBox.Items.Add("✅ נטענו $($groups.Count) קבוצות")
        
        # הצגת פירוט לפי סוג
        $groupsByType = $groups | Group-Object GroupType
        foreach ($groupType in $groupsByType) {
            $resultsBox.Items.Add("   • $($groupType.Name): $($groupType.Count) קבוצות")
        }
    }
    else {
        $resultsBox.Items.Add("❌ לא נמצאו קבוצות")
        $resultsBox.Items.Add("")
        $resultsBox.Items.Add("💡 פתרונות:")
        $resultsBox.Items.Add("   1. Install-Module AzureAD")
        $resultsBox.Items.Add("   2. Connect-AzureAD") 
        $resultsBox.Items.Add("   3. ודא הרשאות לקרוא קבוצות")
    }
    
    foreach ($mb in ($Global:SharedMailboxes | Sort-Object DisplayName)) { 
        $mailboxBox.Items.Add($mb.DisplayName) 
    }
    
    if ($Global:SharedMailboxes.Count -eq 0) {
        $resultsBox.Items.Add("❌ לא נמצאו תיבות דואר משותפות")
        $resultsBox.Items.Add("💡 לחץ 'רענן רשימה' בחלון הראשי")
    }
    else {
        $resultsBox.Items.Add("✅ נטענו $($Global:SharedMailboxes.Count) תיבות דואר משותפות")
    }
    
    $syncBtn.Add_Click({
        if ($groupBox.SelectedItem -and $mailboxBox.SelectedItem) {
            try {
                $resultsBox.Items.Clear()
                $resultsBox.Items.Add("🚀 מתחיל סנכרון...")
                $resultsBox.Refresh()
                
                # זיהוי הקבוצה שנבחרה
                $groupText = $groupBox.SelectedItem.ToString()
                $groupName = $groupText.Split('[')[0].Trim()
                $selectedGroup = $groups | Where-Object { $_.DisplayName -eq $groupName }
                $selectedMailbox = $Global:SharedMailboxes | Where-Object { $_.DisplayName -eq $mailboxBox.SelectedItem }
                
                if (-not $selectedGroup) {
                    $resultsBox.Items.Add("❌ לא נמצאה הקבוצה שנבחרה")
                    return
                }
                
                $resultsBox.Items.Add("📊 קבוצה: $($selectedGroup.DisplayName) [$($selectedGroup.GroupType)]")
                $resultsBox.Items.Add("📮 תיבת יעד: $($selectedMailbox.DisplayName)")
                $resultsBox.Items.Add("🔐 הרשאה: $($permBox.SelectedItem)")
                $resultsBox.Items.Add("")
                $resultsBox.Refresh()
                
                # קבלת חברי הקבוצה
                #$members = Get-GroupMembers -GroupObject $selectedGroup

                $members = Get-MgGroupMember -GroupId $selectedGroup.Identity -All

                if (-not $members -or $members.Count -eq 0) {
                    $resultsBox.Items.Add("❌ לא נמצאו חברים בקבוצה")
                    return
                }
                
                $resultsBox.Items.Add("👥 נמצאו $($members.Count) חברים בקבוצה")
                $resultsBox.Items.Add("🔄 מעבד הרשאות...")
                $resultsBox.Items.Add("")
                $resultsBox.Refresh()
                
                $successCount = 0
                $errorCount = 0
                
                foreach ($member in $members) {
                    try {
                        # זיהוי זהות המשתמש
                        $resultsBox.Items.Add($($member.Id))
                        $resultsBox.Items.Add($member.AdditionalProperties)
                       # $userIdentity = $member.PrimarySmtpAddress ?? $member.UserPrincipalName ?? $member.Mail
                       $userIdentity =($($member.Id))
                       $userDisplayName = $member.DisplayName + $userIdentity
                        $resultsBox.Items.Add("$userIdentity")
                        $resultsBox.Items.Add("$userDisplayName")


                        if (-not $userIdentity) {
                            $resultsBox.Items.Add("⚠️ $userDisplayName - לא נמצאה זהות תקינה")
                            $errorCount++
                            continue
                        }
                        
                        # הוספת הרשאות
                        switch ($permBox.SelectedItem) {
                            "FullAccess" {
                                Add-MailboxPermission -Identity $selectedMailbox.Identity -User $userIdentity -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                            "ReadPermission" {
                                Add-MailboxPermission -Identity $selectedMailbox.Identity -User $userIdentity -AccessRights ReadPermission -InheritanceType All -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                            "SendAs" {
                                Add-RecipientPermission -Identity $selectedMailbox.Identity -Trustee $userIdentity -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
                            }
                        }
                        $resultsBox.Items.Add("✅ $userDisplayName ($userIdentity)")
                        $successCount++
                    }
                    catch {
                        $errorMsg = $_.Exception.Message
                        if ($errorMsg -like "*already exists*" -or $errorMsg -like "*כבר קיימת*") {
                            $resultsBox.Items.Add("ℹ️ $userDisplayName - הרשאה כבר קיימת")
                            $successCount++
                        }
                        else {
                            $resultsBox.Items.Add("❌ $userDisplayName - $errorMsg")
                            $errorCount++
                        }
                    }
                    
                    # רענון תצוגה כל 5 משתמשים
                    if (($successCount + $errorCount) % 5 -eq 0) {
                        $resultsBox.Refresh()
                    }
                }
                
                $resultsBox.Items.Add("")
                $resultsBox.Items.Add("=" * 50)
                $resultsBox.Items.Add("📈 סיכום סנכרון:")
                $resultsBox.Items.Add("✅ הצליחו: $successCount")
                $resultsBox.Items.Add("❌ נכשלו: $errorCount")
                $resultsBox.Items.Add("📊 סה״כ: $($members.Count)")
                $resultsBox.Items.Add("🎉 סנכרון הושלם!")
                
                # הודעת סיכום
                $messageText = "סנכרון הושלם!nn✅ $successCount הרשאות נוספו בהצלחהn❌ $errorCount שגיאותn📊 סה״כ $($members.Count) משתמשים"
                [System.Windows.Forms.MessageBox]::Show($messageText, "סיכום סנכרון", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            }
            catch { 
                $resultsBox.Items.Add("❌ שגיאה כללית: $($_.Exception.Message)")
                [System.Windows.Forms.MessageBox]::Show("שגיאה: $($_.Exception.Message)", "שגיאה", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
        else { [System.Windows.Forms.MessageBox]::Show("בחר קבוצה ותיבת דואר", "שגיאה") }
    })
    
    $syncForm.ShowDialog(); $syncForm.Dispose()
}

# יצירת הטופס הראשי
$form = New-Object System.Windows.Forms.Form
$form.Text = "Exchange Management Tool v2.0"
$form.Size = New-Object System.Drawing.Size(900, 600)
$form.StartPosition = "CenterScreen"
$form.RightToLeft = "Yes"

# כפתורי חיבור
$connectButton = New-Object System.Windows.Forms.Button
$connectButton.Location = "750,20"; $connectButton.Size = "120,35"; $connectButton.Text = "התחבר ל-Exchange"
$connectButton.BackColor = "LightBlue"; $connectButton.Add_Click({ Connect-ToExchange }); $form.Controls.Add($connectButton)

$disconnectButton = New-Object System.Windows.Forms.Button
$disconnectButton.Location = "620,20"; $disconnectButton.Size = "120,35"; $disconnectButton.Text = "נתק"
$disconnectButton.Enabled = $false; $disconnectButton.BackColor = "LightCoral"
$disconnectButton.Add_Click({ Disconnect-Exchange }); $form.Controls.Add($disconnectButton)

# רשימת תיבות דואר
$mailboxLabel = New-Object System.Windows.Forms.Label
$mailboxLabel.Location = "750,70"; $mailboxLabel.Size = "120,20"; $mailboxLabel.Text = "תיבות דואר משותפות:"
$mailboxLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $form.Controls.Add($mailboxLabel)

$mailboxListBox = New-Object System.Windows.Forms.ListBox
$mailboxListBox.Location = "20,90"; $mailboxListBox.Size = "850,300"
$mailboxListBox.Font = New-Object System.Drawing.Font("Consolas", 9); $form.Controls.Add($mailboxListBox)

# כפתורי פעולות
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Location = "750,400"; $refreshButton.Size = "120,30"; $refreshButton.Text = "🔄 רענן רשימה"
$refreshButton.Enabled = $false; $refreshButton.Add_Click({ Load-SharedMailboxes }); $form.Controls.Add($refreshButton)

$createMailboxButton = New-Object System.Windows.Forms.Button
$createMailboxButton.Location = "620,400"; $createMailboxButton.Size = "120,30"; $createMailboxButton.Text = "📧 צור תיבה"
$createMailboxButton.Enabled = $false; $createMailboxButton.BackColor = "LightGreen"
$createMailboxButton.Add_Click({ Create-SharedMailbox }); $form.Controls.Add($createMailboxButton)

$createGroupButton = New-Object System.Windows.Forms.Button
$createGroupButton.Location = "490,400"; $createGroupButton.Size = "120,30"; $createGroupButton.Text = "👥 צור קבוצה"
$createGroupButton.Enabled = $false; $createGroupButton.BackColor = "LightYellow"
$createGroupButton.Add_Click({ Create-Group }); $form.Controls.Add($createGroupButton)

$syncPermissionsButton = New-Object System.Windows.Forms.Button
$syncPermissionsButton.Location = "360,400"; $syncPermissionsButton.Size = "120,30"; $syncPermissionsButton.Text = "🔐 סנכרן הרשאות"
$syncPermissionsButton.Enabled = $false; $syncPermissionsButton.BackColor = "LightSalmon"
$syncPermissionsButton.Add_Click({ Sync-Permissions }); $form.Controls.Add($syncPermissionsButton)

# תווית סטטוס
$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = "20,520"; $statusLabel.Size = "850,30"; $statusLabel.Text = "לא מחובר"
$statusLabel.ForeColor = "Red"; $statusLabel.BorderStyle = "Fixed3D"; $form.Controls.Add($statusLabel)

# הוסף כפתור התחברות ל־Graph לטופס הראשי
$graphButton = New-Object System.Windows.Forms.Button
$graphButton.Location = "490,20"
$graphButton.Size = "120,35"
$graphButton.Text = "התחבר ל־Graph"
$graphButton.BackColor = "LightSkyBlue"
$graphButton.Add_Click({ Connect-ToGraph })
$form.Controls.Add($graphButton)

# הצגת הטופס
[System.Windows.Forms.Application]::Run($form)
