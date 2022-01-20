# requires Connect-ExchangeOnline, Connect-MicrosoftTeams

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN,
  $team
)

$group = (Get-UnifiedGroup -Identity $team)

if ($group) {
  Write-Host 'The group exists.'
} else {
  return Write-Host 'The group does not exist.'
}

$groupId = (Get-UnifiedGroup -Identity $team).ExternalDirectoryObjectId
$team = (Get-Team -GroupId $groupId)

if ($team) {
  Write-Host 'The team exists.'
} else {
  return Write-Host 'The team does not exist.'
}

$teamMembers = (Get-TeamUser -GroupId $groupId).User

if ($teamMembers -contains $UPN) {
  Write-Host 'The user is member of the team.'
} else {
  $toAddToTeam = Read-Host 'The user is not a member of the team. Would you like to add it? [Y/N]'
  if ($toAddToTeam -eq 'Y') {
    Add-TeamUser -GroupId $groupId -User $UPN
    Write-Host 'The user was added to the team.'
  } else {
    return
  }
}
