# התחברות ל-Microsoft Graph
Connect-MgGraph -Scopes "User.Read.All"

# שליפת כל המשתמשים עם הפרטים הדרושים
$users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AccountEnabled,SignInActivity,LastPasswordChangeDateTime,UserType -Filter "userType eq 'Member'"

# יצירת אובייקט עם הנתונים המעוצבים
$userData = @()

foreach ($user in $users) {
    $userObj = [PSCustomObject]@{
        'User ID' = $user.Id
        'Display Name' = $user.DisplayName
        'User Principal Name' = $user.UserPrincipalName
        'Account Enabled' = $user.AccountEnabled
        'Last Sign In' = if ($user.SignInActivity.LastSignInDateTime) { $user.SignInActivity.LastSignInDateTime } else { "Never" }
        'Last Password Change' = if ($user.LastPasswordChangeDateTime) { $user.LastPasswordChangeDateTime } else { "Unknown" }
    }
    $userData += $userObj
}

#$userData |Out-GridView

# ייצוא לקובץ CSV
$userData | Export-Csv -Path "C:\temp\UsersReport.csv" -NoTypeInformation -Encoding UTF8

Write-Host "הדוח נשמר בהצלחה ב: C:\Users\YourUser\Desktop\UsersReport.csv"