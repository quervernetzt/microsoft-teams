# Parameters
[CmdletBinding()]
param (
    [Parameter(
        Mandatory = $false,
        HelpMessage = "Please enter the path to the CSV with members email addresses to add",
        Position = 1)
    ]
    [string]
    $MembersToImportCSVPath = "$PSScriptRoot\members.csv"

    # [Parameter(
    #     Mandatory = $true,
    #     HelpMessage = "Please enter the name of the Teams group",
    #     Position = 2)
    # ]
    # [string]
    # $TeamsGroupName,

    # [Parameter(
    #     Mandatory = $true,
    #     HelpMessage = "Please enter the name of the Private Channel to add the members to",
    #     Position = 3)
    # ]
    # [string]
    # $TeamsGroupPrivateChannelName
)

# Install PS Teams module
# $moduleName = "MicrosoftTeams"
# if (Get-Module -Name $moduleName -ListAvailable) {
#     $toUpdate = Read-Host "Module '$moduleName' is available, do you want to update it? (yes/no)"
#     if ($toUpdate.Trim().ToLower() -eq "yes") {
#         Update-Module $moduleName -Force
#     }
# }
# else {
#     Write-Host "Installing module '$moduleName'..."
#     Install-Module $moduleName -Scope CurrentUser -Force
# }

# Write-Host "Checked module '$moduleName'..."


# # Login
# Connect-MicrosoftTeams

# Get email addresses from CSV
[pscustomobject]$members = Import-Csv -Delimiter "," -Path $MembersToImportCSVPath

foreach ($member in $members) {
    [string]$emailAddress = $member.Email
    Write-Host "Processing member with email address '$emailAddress'..."
}

# Add members to Teams Group (if not already in Teams Group)


# Add members to Private Channel (if not already in the channel)


Write-Host "Done..."
