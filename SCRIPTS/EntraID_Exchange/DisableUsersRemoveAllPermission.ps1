# Import the Active Directory module
#Import-Module ActiveDirectory

# Define the OU and group names
$ouDN = "OU=Account Disabled,OU=Users,OU=MMVP,DC=mmvp,DC=local"
$disabledGroupName = "Disabled Users"
#$domainUsersGroupName = "Domain Users"
#$mvpUsersGroupName = "MVP Users"

# Get the Disabled Users group
$disabledGroup = Get-ADGroup -Identity $disabledGroupName

# Get all users in the specified OU
$users = Get-ADUser -Filter * -SearchBase $ouDN
$users.Count

# Iterate over each user
foreach ($user in $users) {
    Write-Host "Processing user: $($user.Name)"
    
    # Get all groups the user is currently a member of
    $userGroups = Get-ADPrincipalGroupMembership -Identity $user | Where-Object {$_.Name -ne $disabledGroupName}
    
    Write-Host "User is member of $($userGroups.Count) groups (excluding Disabled Users)"
    
    # Add user to Disabled Users group
    try {
        Add-ADGroupMember -Identity $disabledGroup -Members $user -ErrorAction Stop
        Write-Host "Added user to Disabled Users group"
    }
    catch {
        Write-Host "User already member of Disabled Users or error occurred: $($_.Exception.Message)"
    }
    
    $group = Get-ADGroup "Disabled Users" -Properties @("primaryGroupToken")
    
    # Set Disabled Users as the primary group
    try {
        Set-ADUser -Identity $user -Replace @{primaryGroupID = $group.primaryGroupToken}
        Write-Host "Set Disabled Users as primary group"
    }
    catch {
        Write-Host "Error setting primary group: $($_.Exception.Message)"
    }
    
    # Remove user from all other groups (except the primary group which is now Disabled Users)
    foreach ($groupToRemove in $userGroups) {
        try {
            Remove-ADGroupMember -Identity $groupToRemove.Name -Members $user -Confirm:$false -ErrorAction Stop
            Write-Host "Removed user from group: $($groupToRemove.Name)"
        }
        catch {
            Write-Host "Could not remove user from group $($groupToRemove.Name): $($_.Exception.Message)"
        }
    }
    
    # Disable the user account
    try {
        Disable-ADAccount -Identity $user
        Write-Host "Disabled user account: $($user.Name)"
    }
    catch {
        Write-Host "Error disabling account: $($_.Exception.Message)"
    }
    
    Write-Host "Completed processing for user: $($user.Name)"
    Write-Host "----------------------------------------"
}

Write-Host "All users in the OU have been processed: added to Disabled Users, set as primary group, removed from ALL other groups, and disabled."
Start-ADSyncSyncCycle -PolicyType Delta