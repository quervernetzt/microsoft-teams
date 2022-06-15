# Parameters
[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the tenant id")
    ]
    [string]
    $TenantId,

    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the name of the teams group to export from")
    ]
    [string]
    $SourceTeamsGroupName,

    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the file path where to store the list of members")
    ]
    [string]
    $BackupFilePathSourceMembers
)

# Install PS Teams module
[string]$moduleName = "MicrosoftTeams"
if (Get-Module -Name $moduleName -ListAvailable) {
    [string]$toUpdate = Read-Host "Module '$moduleName' is available, do you want to update it? (yes/no)"
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
Connect-MicrosoftTeams -TenantId $TenantId
Connect-AzAccount -TenantId $TenantId

# Get members of source
[string]$sourceTeamsGroupId = (Get-Team -DisplayName $SourceTeamsGroupName | Where-Object { $_.DisplayName -eq $SourceTeamsGroupName }).GroupId
[object[]]$sourceTeamsGroupMembers = Get-TeamUser -GroupId $sourceTeamsGroupId

[int]$userCount = ($sourceTeamsGroupMembers | Measure-Object).Count
Write-Output "Retrieved '$userCount' users."

# Get additional information for users
[System.Collections.ArrayList]$users= @()
foreach($sourceTeamsGroupMember in $sourceTeamsGroupMembers) {
    [string]$userId = $sourceTeamsGroupMember.UserId
    $user = (Get-AzADUser -ObjectId $userId -Select @('City', 'UsageLocation') -AppendSelected) | Select-Object -Property Id, DisplayName, Mail, UserPrincipalName, UsageLocation, City
    $users.Add($user)
}

# Backup list of source members to csv
$users | Export-Csv -Path $BackupFilePathSourceMembers -NoTypeInformation
Write-Host "Saved members of source Teams group '$SourceTeamsGroupName' in '$BackupFilePathSourceMembers'..."

Write-Host "Done..."