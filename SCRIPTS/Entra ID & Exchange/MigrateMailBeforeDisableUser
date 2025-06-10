# בדיקה אם המודול מותקן
$moduleName = "ExchangeOnlineManagement"
$installed = Get-Module -ListAvailable -Name $moduleName

if (-not $installed) {
    Write-Host "$moduleName לא מותקן. מתקין..." -ForegroundColor Yellow

    try {
        Install-Module -Name $moduleName -Scope AllUsers -Force -AllowClobber
        Write-Host "$moduleName הותקן בהצלחה." -ForegroundColor Green
        
        # ניסיון לייבא את המודול
        try {
            Import-Module $moduleName -Force
            Write-Host "המודול יובא בהצלחה." -ForegroundColor Green
        } catch {
            Write-Error "שגיאה בייבוא המודול: $_"
            exit
        }
    } catch {
        Write-Error "ההתקנה נכשלה: $_"
        exit
    }
} else {
    Write-Host "$moduleName כבר מותקן." -ForegroundColor Cyan
}



# Load Windows Forms
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Add Users to BlockEmailRecipients Group"
$form.Size = New-Object System.Drawing.Size(500,400)
$form.StartPosition = "CenterScreen"

# Create TextBox for users input
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.Size = New-Object System.Drawing.Size(450,250)
$textBox.Location = New-Object System.Drawing.Point(20,20)
$textBox.ScrollBars = "Vertical"
$textBox.AcceptsReturn = $true
$form.Controls.Add($textBox)

# Create Button
$button = New-Object System.Windows.Forms.Button
$button.Text = "Start Processing"
$button.Size = New-Object System.Drawing.Size(150,40)
$button.Location = New-Object System.Drawing.Point(170,290)
$form.Controls.Add($button)

# Button Click Action
$button.Add_Click({
    $UPNUsers = $textBox.Lines | Where-Object { $_.Trim() -ne "" }

    if ($UPNUsers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter at least one user UPN.",
            "Input Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $form.Close()

    # Connect to Exchange Online
    Connect-ExchangeOnline

    foreach ($UPNUser in $UPNUsers) {
        try {
            Write-Host "Processing $UPNUser..." -ForegroundColor Cyan

            # Set mailbox to shared type
            Set-Mailbox "$UPNUser" -Type Shared -ErrorAction Stop

            # Output mailbox details
            Get-Mailbox -Identity "$UPNUser" | Format-Table Alias, Name, RecipientTypeDetails

            # Add user to the Office 365 group
            Add-UnifiedGroupLinks -Identity "BlockEmailRecipients@mmvp.co.il" -LinkType "Members" -Links "$UPNUser" -ErrorAction Stop

            # Validate if the user is in the group
            $groupMembers = Get-UnifiedGroupLinks -Identity "BlockEmailRecipients@mmvp.co.il" -LinkType "Members"
            $userInGroup = $groupMembers | Where-Object { $_.Alias -eq $UPNUser.Split('@')[0] }

            if ($userInGroup) {
                Write-Host "$UPNUser has been successfully added to the BlockEmailRecipients group." -ForegroundColor Green
            } else {
                Write-Host "$UPNUser is NOT in the BlockEmailRecipients group." -ForegroundColor Red
            }
        } catch {
            Write-Host "An error occurred while processing UPNUser: $($_)" -ForegroundColor Red
        }

    }

    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Confirm:$false

    <#  
    # Show final message
    [System.Windows.Forms.MessageBox]::Show(
        "Finished processing all users!",
        "Completed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    #>
})

# Show the form
$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
