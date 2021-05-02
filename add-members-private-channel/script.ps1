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
[string]$moduleName = "MicrosoftTeams"
[string]$validatedVersion = "2.2.0"

if (Get-Module -Name $moduleName -ListAvailable) {
    Write-Host "Microsoft Teams module available..."
    [psmoduleinfo]$module = Get-Module -Name $moduleName -ListAvailable
    [version]$moduleVersion = $module.Version

    if ($moduleVersion -eq $validatedVersion) {
        Write-Host "Validated version installed..."
    }
    else {
        throw "Please check the version installed to support the required commands..."
    }
}
else {
    Write-Host "Installing Microsoft Teams module..."
    Install-Module PowerShellGet -Force -AllowClobber
    Install-Module -Name MicrosoftTeams -AllowPrerelease -RequiredVersion 2.2.0-preview
}

Write-Host "Checked module '$moduleName' with version '$validatedVersion'..."

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
[int]$numberOfMembersToAddToTeamsGroup = ($teamsGroupMembersTarget | Measure-Object).Count

if ($numberOfMembersToAddToTeamsGroup -gt 0) {
    $teamsGroupMembersTarget | ForEach-Object {
        [string]$emailAddress = $_.Email
    
        Add-TeamUser -GroupId $teamsGroupId -User $emailAddress -Role Member
        Write-Host "Added user '$emailAddress' with role 'Member' to group '$TeamsGroupName'..."
    }
    
    Write-Host "Added users to Teams Group..."
}
else {
    Write-Host "No users to add to Teams Group..."
}


########################
# Get current users of Private Channel
########################
[object]$privateChannelMembersCurrent = Get-TeamChannelUser -GroupId $teamsGroupId -DisplayName $TeamsGroupPrivateChannelName -Role Member
[object]$privateChannelOwnersCurrent = Get-TeamChannelUser -GroupId $teamsGroupId -DisplayName $TeamsGroupPrivateChannelName -Role Owner

########################
# Add members to Private Channel (if not already in the channel)
########################
[object[]]$privateChannelMembersTarget = $membersToAdd | Where-Object { ($_.Email -notin $privateChannelMembersCurrent.User) -and ($_.Email -notin $privateChannelOwnersCurrent.User) }
[int]$numberOfMembersToAddToPrivateChannel = ($privateChannelMembersTarget | Measure-Object).Count

if ($numberOfMembersToAddToPrivateChannel -gt 0) {
    $privateChannelMembersTarget | ForEach-Object {
        [string]$emailAddress = $_.Email
    
        Add-TeamChannelUser -GroupId $teamsGroupId -DisplayName $TeamsGroupPrivateChannelName -User $emailAddress
        Write-Host "Added user '$emailAddress' with role 'Member' to Private Channel '$TeamsGroupPrivateChannelName'..."
    }
    
    Write-Host "Added users to Private Channel..."
}
else {
    Write-Host "No users to add to Private Channel..."
}

# ########################
# Validation
# ########################


Write-Host "Done..."
