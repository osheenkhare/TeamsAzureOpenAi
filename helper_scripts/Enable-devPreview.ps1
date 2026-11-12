Connect-MicrosoftTeams
# Force Public Preview (R3.6) for everyone on the Global policy
Set-CsTeamsUpdateManagementPolicy -Identity Global -AllowPublicPreview Forced