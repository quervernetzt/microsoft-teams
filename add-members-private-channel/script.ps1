########################
# Parameters
########################
[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Please enter the path to the CSV with members email addresses to add")
    ]
    [string]
    $MembersToImportCSVPath = "$PSScriptRoot\members.csv",

    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the name of the Teams group")
    ]
    [string]
    $TeamsGroupName,

    [Parameter(
        Mandatory = $true,
        HelpMessage = "Please enter the name of the Private Channel to add the members to")
    ]
    [string]
    $TeamsGroupPrivateChannelName
)

########################
# Check teams module
########################
$moduleName = "MicrosoftTeams"
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

########################
# Login
########################
Connect-MicrosoftTeams

########################
# Get email addresses from CSV
########################
[pscustomobject]$membersToAdd = Import-Csv -Delimiter "," -Path $MembersToImportCSVPath

foreach ($memberToAdd in $membersToAdd) {
    [string]$emailAddress = $memberToAdd.Email
    Write-Host "Processing member with email address '$emailAddress'..."
}

########################
# Get current users of Teams group
########################
[string]$teamsGroupId = (Get-Team -DisplayName $TeamsGroupName).GroupId

[object]$teamsGroupMembersCurrent = Get-TeamUser -GroupId $teamsGroupId -Role Member
[object]$teamsGroupOwnersCurrent = Get-TeamUser -GroupId $teamsGroupId -Role Owner

########################
# Add members to Teams Group (if not already in Teams Group)
########################
[object[]]$teamsGroupMembersTarget = $membersToAdd | Where-Object { ($_.Email -notin $teamsGroupMembersCurrent.User) -and ($_.Email -notin $teamsGroupOwnersCurrent.User) }

$teamsGroupMembersTarget | ForEach-Object {
    [string]$emailAddress = $_.Email

    Add-TeamUser -GroupId $teamsGroupId -User $emailAddress -Role Member
    Write-Host "Added user '$emailAddress' with role 'Member' to group '$TeamsGroupName'..."
}

Write-Host "Finished adding users, validating..."

########################
# Add members to Private Channel (if not already in the channel)
########################

Write-Host "Done..."
