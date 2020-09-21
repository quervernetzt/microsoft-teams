# Parameters
$sourceTeamsGroupName = "xxx"
$targetTeamsGroupName = "xxx"
$deleteUsersFromSourceTeamsGroup = $true
$backupFilePathSourceMembers = "C:\xxx\migrate-members\Teams_Source_Members_Backup.csv"
$backupFilePathTargetMembers = "C:\xxx\migrate-members\Teams_Target_Members_Backup.csv"
$backupFilePathTargetOwners = "C:\xxx\migrate-members\Teams_Target_Owners_Backup.csv"


# Install PS Teams module
$moduleName = "MicrosoftTeams"
if (Get-Module -Name $moduleName -ListAvailable) {
    $toUpdate = Read-Host "Module '$moduleName' is available, do you want to update it? (yes/no)"
    if ($toUpdate.Trim().ToLower() -eq "yes") {
        Update-Module $moduleName -Force
    }
} else {
    Write-Host "Installing module '$moduleName'..."
    Install-Module $moduleName -Scope CurrentUser -Force
}

Write-Host "Checked module '$moduleName'..."


# Login
Connect-MicrosoftTeams


# Get members of source
$sourceTeamsGroupId = (Get-Team -DisplayName $sourceTeamsGroupName).GroupId
$sourceTeamsGroupMembers = Get-TeamUser -GroupId $sourceTeamsGroupId -Role Member


# Backup list of source members to csv
$sourceTeamsGroupMembers | Export-Csv -Path $backupFilePathSourceMembers -NoTypeInformation
Write-Host "Retrieved and backuped members of source Teams group '$sourceTeamsGroupName'..."


# Add users of target group
$targetTeamsGroupId = (Get-Team -DisplayName $targetTeamsGroupName).GroupId

$targetTeamsGroupMembersCurrent = Get-TeamUser -GroupId $targetTeamsGroupId -Role Member
$targetTeamsGroupOwnersCurrent = Get-TeamUser -GroupId $targetTeamsGroupId -Role Owner


# Backup list of target members and owners to csv
$targetTeamsGroupMembersCurrent | Export-Csv -Path $backupFilePathTargetMembers -NoTypeInformation
Write-Host "Retrieved and backuped members of target Teams group '$targetTeamsGroupName'..."

$targetTeamsGroupOwnersCurrent | Export-Csv -Path $backupFilePathTargetOwners -NoTypeInformation
Write-Host "Retrieved and backuped owners of target Teams group '$targetTeamsGroupName'..."


# Add members to target group
$targetTeamsGroupMembersTarget = $sourceTeamsGroupMembers | Where-Object { ($_.UserId -notin $targetTeamsGroupMembersCurrent.UserId) -and ($_.UserId -notin $targetTeamsGroupOwnersCurrent.UserId) }

$targetTeamsGroupMembersTarget | ForEach-Object {
    $userId = $_.UserId
    $userName = $_.User

    Add-TeamUser -GroupId $targetTeamsGroupId -User $userId -Role Member
    Write-Host "Added user '$userName' with role 'Member' to group '$targetTeamsGroupName'..."
}

Write-Host "Finished adding users, validating..."


# Validate
$targetTeamsGroupMembersCheck = Get-TeamUser -GroupId $targetTeamsGroupId -Role Member
$targetTeamsGroupOwnersCheck = Get-TeamUser -GroupId $targetTeamsGroupId -Role Owner
$targetTeamsGroupMembersDifference = $sourceTeamsGroupMembers | Where-Object { ($_.UserId -notin $targetTeamsGroupMembersCheck.UserId) -and ($_.UserId -notin $targetTeamsGroupOwnersCheck.UserId) }

if (($targetTeamsGroupMembersDifference | Measure-Object).Count -eq 0) {
    Write-Host "Validation successful..."
} else {
    Write-Host "Validation failed for: "
    $targetTeamsGroupMembersDifference | ForEach-Object {
        Write-Host $_.User
    }
    Write-Host ""
    throw "Validation failed..."
}


# Remove users from source group (if selected)
if ($deleteUsersFromSourceTeamsGroup) {
    $sourceTeamsGroupMembers | ForEach-Object {
        $userId = $_.UserId
        $userName = $_.User
        Remove-TeamUser -GroupId $sourceTeamsGroupId -User $userId
        Write-Host "Removed user '$userName' from group '$sourceTeamsGroupName'..."
    } 
}

Write-Host "Done..."