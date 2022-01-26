# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Author: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Dependencies: Connect-MsolService, Connect-ExchangeOnline, Connect-MicrosoftTeams
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Usage: 
# Connect-MsolService
# Connect-ExchangeOnline
# Connect-MicrosoftTeams
# .\tms-CannotAddUserToPrivateChannel.ps1 user@domain.com team@domain.com
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN,
  $groupSMTP
)

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# FUNCTIONS
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
function Is-UserLicensed {
  param (
    [string[]]$UserPrincipalName
  )

  $isLicensed = (Get-MsolUser -UserPrincipalName $UserPrincipalName).isLicensed

  if ($isLicensed) {
    return $true
  } else {
    return $false
  }
}

# function Check-OfficeLicense { # ERRORS OUT ??
#   param (
#     [string[]]$UserPrincipalName
#   )
  
#   $UserPrincipalName = $UserPrincipalName | Out-String

#   Write-Host $UserPrincipalName
#   Write-Host $UserPrincipalName.gettype()

#   $SKUs = (Get-MsolUser -UserPrincipalName $UserPrincipalName).Licenses.AccountSkuId #.split(":")[1]
#   $ServicePlans = (Get-MsolUser -UserPrincipalName $UserPrincipalName).Licenses.ServiceStatus.ServicePlan.ServiceName

#   $licenses = @{
#     SKU = @()
#     ServicePlans = ''
#   }

#   if ($SKUs.length -gt 1) {
#     foreach ($SKU in $SKUs) {
#       $licenses.SKU += $SKU.split(":")[1]
#     }
#   } else {
#     $licenses.SKU += $SKUs.split(":")[1]
#   }

#   $licenses.ServicePlans = $ServicePlans

#   return $licenses
# }

Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'User:'$UPN
Write-Host 'Team:'$groupSMTP
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'

$user = (Get-MsolUser -UserPrincipalName $UPN)

Write-Host 'Checking if the user exists:'

if ($user) {
  Write-Host -ForegroundColor Green 'The user exists.'
} else {
  return Write-Host -ForegroundColor Red 'The user does not exist.'
}

$group = (Get-UnifiedGroup -Identity $groupSMTP)

Write-Host 'Checking if the M365 Group exists:'

if ($group) {
  Write-Host -ForegroundColor Green 'The group exists.'
} else {
  return Write-Host -ForegroundColor Red 'The group does not exist.'
}

$groupId = (Get-UnifiedGroup -Identity $groupSMTP).ExternalDirectoryObjectId
$team = (Get-Team -GroupId $groupId)


Write-Host 'Checking if a team is exists for the provided M365 Group:'
if ($team) {
  Write-Host -ForegroundColor Green 'The team exists.'
} else {
  return Write-Host -ForegroundColor Red 'The team does not exist.'
}

$teamMembers = (Get-TeamUser -GroupId $groupId).User

Write-Host 'Checking if the user is a member of the team:'

if ($teamMembers -contains $UPN) {
  Write-Host -ForegroundColor Green 'The user is member of the team.'
} else {
  $toAddToTeam = Read-Host -ForegroundColor Orange 'The user is not a member of the team. Would you like to add it? [Y/N]'
  if ($toAddToTeam -eq 'Y') {
    Add-TeamUser -GroupId $groupId -User $UPN
    Write-Host -ForegroundColor Green 'The user was added to the team.'
  } else {
    return
  }
}

# $sub = (Check-OfficeLicense $UPN).SKU
# $sps = (Check-OfficeLicense $UPN).ServicePlans

$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\CannotAddUserToPrivateChannel.txt"
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User $($UPN) (TMS):"
  Get-CsOnlineUser $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User $($UPN) (MSO):"
  Get-MsolUser -UserPrincipalName $UPN
  # Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  # Write-Host "User $($UPN) subscriptions and service plans (MSO):"
  # Write-Host 'SKUs:'
  # Write-Host $sub
  # Write-Host 'Service plans:'
  # Write-Host $sps
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Group $($GROUP) (EXO):"
  Get-UnifiedGroup -Identity $groupSMTP
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Team $($GROUP) (TMS):"
  Get-Team -GroupId $groupId
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Team $($GROUP) membership (TMS):"
  Get-TeamUser -GroupId $groupId
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Stop-Transcript
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
