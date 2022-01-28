# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# DEPENDENCIES: Connect-AzureAD
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USAGE: 
# Connect-AzureAD
# .\tms-GuestCannotSignIn.ps1 user@domain.com
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Back it up with https://aka.ms/TeamsGuestAccessDiag
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TODO: Guest user count (MAU Billing Model) ??

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN
)

Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'Guest user:'$UPN
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# FUNCTIONS
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
function Get-GuestExtUPN {
  param (
    [string]$UserPrincipalName,
    [string]$CompanyDomain
  )

# user@domain.com -> user_domain.com#EXT#@domain.onmicrosoft.com
  return $UPN.replace('@', '_')  + '#EXT#' + '@' + $CompanyDomain
}

function Get-UserTenantId {
  param (
    [string]$UserPrincipalName
  )

  $tenantId = ''

  $userDomain = $UPN.split('@')[1]

  $url = "https://login.windows.net/$($userDomain)/.well-known/openid-configuration"

  try {
    $tenantId = (curl $url | ConvertFrom-Json).token_endpoint.split('/')[3]
  } catch {}

  return $tenantId
}
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# CONVERTING TO #EXT# UPN
$MSARealm = '9cd80435-793b-4f48-844b-6b3f37d1c1f3' # homing all MSA domains (gmail.com, outlook.com, live.com...)
$verifiedDomains = (Get-AzureADTenantDetail).VerifiedDomains
$initialDomain = ''

foreach ($vd in $verifiedDomains) {
  if ($vd.Initial) {
    $initialDomain = $vd.Name
  }
}

$guestExtUPN = Get-GuestExtUPN $UPN $initialDomain
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# GUEST USER
$guest = ''

try {
  $guest = (Get-AzureADUser -ObjectId $guestExtUPN)
} catch {} # exception if guest user does not exist (AAD module)


Write-Host 'Checking if the guest user exists:'
if ($guest) {
  Write-Host -ForegroundColor Green 'The guest user exists in AAD.'
} else {
  return
}

# TYPE
$userType = $guest.userType

Write-Host 'Checking if the guest user is of correct type:'
if ($userType -eq 'Guest') {
  Write-Host -ForegroundColor Green 'The guest user is of correct type:'$userType
} else {
  return Write-Host -ForegroundColor Red 'The guest user is of incorrect type:'$userType
}

# INVITATION STATE
$userState = $guest.userState

Write-Host 'Checking if the guest user redeemed the invitation:'
if ($userState -eq 'Accepted') {
  Write-Host -ForegroundColor Green 'The guest user redeemed the invitation.'
} elseif ($userState -eq 'PendingAcceptance') {
  return Write-Host -ForegroundColor Red 'The guest user did not redeem the invitation.'
}

# HOMING (MSA vs AAD)
$guestTID = Get-UserTenantId $UPN

# possible flaw: MSA user with custom domain alias added before domain was added to an AAD -> would result in AAD source instead of MSA
Write-Host 'Checking the guest user state:'
if ($guestTID -eq $MSARealm) {
  Write-Host 'The guest user is of state 2 (Microsoft Services Account).'
} else {
  Write-Host 'The guest user is of state 1 (Azure Active Directory).'
}
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\GuestCannotSignIn.txt"
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Guest user:'$UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Guest user (AAD):"
  Get-AzureADUser -ObjectId $GuestExtUPN | fl
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Guest user homing (tenantId):'$guestTID
  Stop-Transcript
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
