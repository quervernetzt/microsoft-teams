# Parameters
[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the name of the teams group to export from",
        Position = 1)
    ]
    [string]
    $SourceTeamsGroupName,

    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the name of the teams group to export to",
        Position = 2)
    ]
    [string]
    $TargetTeamsGroupName,

    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter boolean to indicate whether to delete the exported users from the source teams group",
        Position = 3)
    ]
    [bool]
    $DeleteUsersFromSourceTeamsGroup,

    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the path where to save the csv with the members from the source Teams group to export",
        Position = 4)
    ]
    [string]
    $BackupFilePathSourceMembers,

    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the path where to save the csv with the members of the target Teams group",
        Position = 5)
    ]
    [string]
    $BackupFilePathTargetMembers,

    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the path where to save the csv with the owners of the target Teams group",
        Position = 6)
    ]
    [string]
    $BackupFilePathTargetOwners
)

# Install PS Teams module
$moduleName = "MicrosoftTeams"
if (Get-Module -Name $moduleName -ListAvailable) {
    $toUpdate = Read-Host "Module '$moduleName' is available, do you want to update it? (yes/no)"
    if ($toUpdate.Trim().ToLower() -eq "yes") {
        Update-Module $moduleName -Force
    }
}
else {
    Write-Host "Installing module '$moduleName'..."
    Install-Module $moduleName -Scope CurrentUser -Force
}

Write-Host "Checked module '$moduleName'..."


# Login
Connect-MicrosoftTeams


# Get members of source
$sourceTeamsGroupId = (Get-Team -DisplayName $SourceTeamsGroupName).GroupId
$sourceTeamsGroupMembers = Get-TeamUser -GroupId $sourceTeamsGroupId -Role Member


# Backup list of source members to csv
$sourceTeamsGroupMembers | Export-Csv -Path $BackupFilePathSourceMembers -NoTypeInformation
Write-Host "Retrieved and backuped members of source Teams group '$SourceTeamsGroupName'..."


# Get users of target group
$targetTeamsGroupId = (Get-Team -DisplayName $TargetTeamsGroupName).GroupId

$targetTeamsGroupMembersCurrent = Get-TeamUser -GroupId $targetTeamsGroupId -Role Member
$targetTeamsGroupOwnersCurrent = Get-TeamUser -GroupId $targetTeamsGroupId -Role Owner


# Backup list of target members and owners to csv
$targetTeamsGroupMembersCurrent | Export-Csv -Path $BackupFilePathTargetMembers -NoTypeInformation
Write-Host "Retrieved and backuped members of target Teams group '$TargetTeamsGroupName'..."

$targetTeamsGroupOwnersCurrent | Export-Csv -Path $BackupFilePathTargetOwners -NoTypeInformation
Write-Host "Retrieved and backuped owners of target Teams group '$TargetTeamsGroupName'..."


# Add members to target group
$targetTeamsGroupMembersTarget = $sourceTeamsGroupMembers | Where-Object { ($_.UserId -notin $targetTeamsGroupMembersCurrent.UserId) -and ($_.UserId -notin $targetTeamsGroupOwnersCurrent.UserId) }

$targetTeamsGroupMembersTarget | ForEach-Object {
    $userId = $_.UserId
    $userName = $_.User

    Add-TeamUser -GroupId $targetTeamsGroupId -User $userId -Role Member
    Write-Host "Added user '$userName' with role 'Member' to group '$TargetTeamsGroupName'..."
}

Write-Host "Finished adding users, validating..."


# Validate
$targetTeamsGroupMembersCheck = Get-TeamUser -GroupId $targetTeamsGroupId -Role Member
$targetTeamsGroupOwnersCheck = Get-TeamUser -GroupId $targetTeamsGroupId -Role Owner
$targetTeamsGroupMembersDifference = $sourceTeamsGroupMembers | Where-Object { ($_.UserId -notin $targetTeamsGroupMembersCheck.UserId) -and ($_.UserId -notin $targetTeamsGroupOwnersCheck.UserId) }

if (($targetTeamsGroupMembersDifference | Measure-Object).Count -eq 0) {
    Write-Host "Validation successful..."
}
else {
    Write-Host "Validation failed for: "
    $targetTeamsGroupMembersDifference | ForEach-Object {
        Write-Host $_.User
    }
    Write-Host ""
    throw "Validation failed..."
}


# Remove users from source group (if selected)
if ($DeleteUsersFromSourceTeamsGroup) {
    $sourceTeamsGroupMembers | ForEach-Object {
        $userId = $_.UserId
        $userName = $_.User
        Remove-TeamUser -GroupId $sourceTeamsGroupId -User $userId
        Write-Host "Removed user '$userName' from group '$SourceTeamsGroupName'..."
    } 
}

Write-Host "Done..."
