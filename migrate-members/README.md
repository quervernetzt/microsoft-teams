# General

- Goal is to export the members of a Teams group and to import them to another Teams group

# How to use it

- Execute the script providing the required parameters

- Example

`.\migrate-members\script.ps1 -SourceTeamsGroupName "Group1" -TargetTeamsGroupName "Group2" -DeleteUsersFromSourceTeamsGroup $true -BackupFilePathSourceMembers ".\migrate-members\Teams_Source_Members_Backup.csv" -BackupFilePathTargetMembers ".\migrate-members\Teams_Target_Members_Backup.csv" -BackupFilePathTargetOwners ".\migrate-members\Teams_Target_Owners_Backup.csv"`


# Resources

[1] [Microsoft Teams PowerShell Overview](https://docs.microsoft.com/en-us/microsoftteams/teams-powershell-overview)

[2] [MicrosoftTeamsPowerShell](https://docs.microsoft.com/en-us/powershell/module/teams/?view=teams-ps)