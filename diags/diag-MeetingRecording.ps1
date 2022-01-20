## Requires Connect-MicrosoftTeams

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN
)

$policyName = (Get-CsOnlineUser $UPN).TeamsMeetingPolicy

Write-Host "UPN:"$UPN
Write-Host "Meeting policy:"$policyName

if (!$policyName)
{
  $policyName = 'Global'
}
$meetingPolicy = (Get-CsTeamsMeetingPolicy -Identity $policyName)

if ($meetingPolicy.AllowCloudRecording -eq $true) {
  Write-Host "`nAllowCloudRecording: The user's meeting policy allows for cloud recording."
} else {
  $toEnable = Read-Host "`nAllowCloudRecording: The user's meeting policy does not for cloud recording. Would you like to enable it? [Y/N]"
  
  if ($toEnable -eq 'Y') {
    Set-CsTeamsMeetingPolicy -Identity $policyName -AllowCloudRecording $true
    Write-Host "AllowCloudRecording: The user's meeting policy now allows for cloud recording."
  } else {
    return 
  }
}

if ($meetingPolicy.RecordingStorageMode -eq 'Stream') {
  $toChangeToODSP = Read-Host "RecordingStorageMode: The user's meeting recordings are stored in Stream. Would you like to change to OneDrive/SharePoint? [Y/N]"
  if ($toChangeToODSP -eq 'Y') {
    Set-CsTeamsMeetingPolicy -Identity $policyName -RecordingStorageMode "OneDriveForBusiness"
    Write-Host "RecordingStorageMode: The user's meeting recordings are now stored in OneDrive/SharePoint."
  } else {
    return
  }
} else {
  Write-Host "RecordingStorageMode: The user's meeting recordings are stored in OneDrive/SharePoint."
}
