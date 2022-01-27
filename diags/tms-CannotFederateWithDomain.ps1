# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# DEPENDENCIES: Connect-MsolService, Connect-MicrosoftTeams
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USAGE: 
# Connect-MsolService
# Connect-MicrosoftTeams
# .\tms-CannotFederateWithDomain.ps1 user@domain.com domain.com
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Back it up with https://aka.ms/TeamsFederationDiag
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN,
  $Domain
)

# AVOID TRUNCATING LISTS AT 4TH ITEM
$FormatEnumerationLimit = -1

Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'User:'$UPN
Write-Host 'External domain:'$Domain
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# FUNCTIONS
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
function Get-OfficeUserLicense {
  param (
    [string]$UserPrincipalName
  )

  $SKUs = (Get-MsolUser -UserPrincipalName $UserPrincipalName).Licenses.AccountSkuId
  $ServicePlans = (Get-MsolUser -UserPrincipalName $UserPrincipalName).Licenses.ServiceStatus.ServicePlan.ServiceName

  $licenses = @{
    isLicensed = (Get-MsolUser -UserPrincipalName $UserPrincipalName).isLicensed
    SKU = @()
    ServicePlans = ''
  }

  if ($SKUs.length -gt 1) {
    foreach ($SKU in $SKUs) {
      $licenses.SKU += $SKU.split(":")[1]
    }
  } else {
    try {
      $licenses.SKU += $SKUs.split(":")[1]
    } catch {} # exception if user is unlicensed
  }

  $licenses.ServicePlans = $ServicePlans

  return $licenses
}

function Get-TenantFederationConfig {
  return Get-CsTenantFederationConfiguration
}

function Get-AllowList {
  return (Get-CsTenantFederationConfiguration).AllowedDomains
}

function Get-IsAllowList {
  param (
    $AllowedDomains
  )

  $AllowedDomains = $AllowedDomains.Element | Out-String

  $AllowedDomainsElement = $AllowedDomains.split("<")[1].split(" ")[0]

  if ($AllowedDomainsElement -eq 'AllowAllKnownDomains') {
    return $false
  } elseif ($AllowedDomainsElement -eq 'AllowList') {
    return $true
  }
}

function Get-ParsedAllowList {
  param (
    $AllowedDomains
  )

  $allowList = @()

  foreach ($domain in $AllowedDomains.AllowedDomain) {
    $domain = $domain | Out-String
    $allowList += $domain.split(":")[1].trim()
  }

  return $allowList
}

function Set-AllowList {
  param (
    $parsedAL,
    $dmn
  )

  $aList = @()

  foreach ($d in $parsedAL) {
    $aList += $d
  }

  $aList += $dmn
  $list = New-Object Collections.Generic.List[String]

  foreach ($i in $aList) {
    $list.add($i)
  }

  Set-CsTenantFederationConfiguration -AllowedDomainsAsAList $list

  Write-Host -ForegroundColor Green 'The provided domain was added to the allow list.'
}

function Get-BlockList {
  return (Get-CsTenantFederationConfiguration).BlockedDomains
}

function Get-IsBlockList {
  param (
    $BlockedDomains
  )

  if ($BlockedDomains -eq $null) {
    return $false
  } else {
    return $true
  }
}

function Get-ParsedBlocklist {
  param (
    $BlockedDomains
  )

  $blockList = @()

  foreach ($domain in $BlockedDomains) {
    $domain = $domain | Out-String
    $blockList += $domain.split(":")[1].trim()
  }
  return $blockList
}

function Set-BlockList {
  param (
    $parsedBL,
    $dmn
  )

  $bList = @()

  foreach ($d in $parsedBL) {
    if ($d -ne $dmn) {
      $bList += $d
    }
  }

  Set-CsTenantFederationConfiguration -BlockedDomains $null

  foreach ($i in $bList) {
    $x = New-CsEdgeDomainPattern -Domain $i
    Set-CsTenantFederationConfiguration -BlockedDomains @{Add=$x}
  }

  Write-Host -ForegroundColor Green 'The provided domain was removed from the block list.'
}

function Get-ExternalAccessPolicyName {
  param (
    [string]$UserPrincipalName
  )

  $pN = (Get-CsOnlineUser $UserPrincipalName).ExternalAccessPolicy

   if ($pN -eq $null) {
    $pN = 'Global'
  }

  return $pN
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USER VALIDATION
$user = (Get-MsolUser -UserPrincipalName $UPN)

Write-Host 'Checking if the user exists:'
if ($user) {
  Write-Host -ForegroundColor Green 'The user exists.'
} else {
  return Write-Host -ForegroundColor Red 'The user does not exist.'
}

$userLicense = Get-OfficeUserLicense $UPN

Write-Host 'Checking if the user is licensed:'
if ($userLicense.isLicensed) {
  Write-Host -ForegroundColor Green 'The user is licensed.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed.'
}

Write-Host 'Checking if the user is licensed for Teams and Skype for Business Online:'
if ($userLicense.ServicePlans -contains 'TEAMS1') {
  Write-Host -ForegroundColor Green 'The user is licensed for Teams.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Teams.'
}

if ($userLicense.ServicePlans -contains 'MCOSTANDARD' -OR $userLicense.ServicePlans -contains 'MCO_TEAMS_IW') {
  Write-Host -ForegroundColor Green 'The user is licensed for Skype for Business Online.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Skype for Business Online.'
}

$tmsUser = (Get-CsOnlineUser $UPN)

Write-Host 'Checking if the user is SIP enabled:'
if ($tmsUser.Enabled) {
  Write-Host -ForegroundColor Green 'The user is SIP enabled.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not SIP enabled.'
}

Write-Host 'Checking if the user has any MCOValidationError:'
if (!$tmsUser.MCOValidationError) {
  Write-Host -ForegroundColor Green 'The user does not have any MCOValidationError.'
} else {
  return Write-Host -ForegroundColor Red 'The user has an MCOValidationError:'$tmsUser.MCOValidationError
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USER EXTRNAL ACCESS POLICY
$policy = $tmsUser.ExternalAccessPolicy

if ($policy -eq $null) {
  $policy = 'Global'
}

Write-Host 'Checking if the user has been granted an external access policy that allows to communicate with external users:'
if ($policy -eq 'FederationAndPICDefault' -OR $policy -eq 'FederationOnly') {
  Write-Host -ForegroundColor Green 'The user has been granted an external access policy that allows to communicate with external users.'
} elseif ($policy -eq 'NoFederationAndPIC') {
  Write-Host -ForegroundColor Red 'The user has been granted an external access policy that does not allow to communicate with external users.'
  $toChangeEAPolicy = Read-Host 'Would you like to change it? [Y/N]'
  if ($toChangeEAPolicy -eq 'Y') {
    Grant-CsExternalAccessPolicy -Identity $UPN -PolicyName FederationAndPICDefault
    Write-Host -ForegroundColor Green 'The user has now been granted an external access policy that allows to communicate with external users.'
  } else {
    return
  }
} else {
  $enableFederationAccess = (Get-CsExternalAccessPolicy -Identity $policy).enableFederationAccess
  if ($enableFederationAccess) {
    Write-Host -ForegroundColor Green 'The user has been granted an external access policy that allows to communicate with external users.'
  } else {
    Write-Host -ForegroundColor Red 'The user been granted an external access policy that does not allow to communicate with external users.'
    $toEnableEAOnPolicy = Read-Host 'Would you like to enable federation on the policy? [Y/N]'
    if ($toEnableEAOnPolicy -eq 'Y') {
      Set-CsExternalAccessPolicy -Identity $policy -EnableFederationAccess $true
      Write-Host -ForegroundColor Green 'The external access policy granted to the user now allows to communicate with external users.'
    } else {
      return
    }
  }
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# TENANT FEDERATION
$federationConfig = Get-TenantFederationConfig

Write-Host 'Checking if the organization allows to communicate with Teams and Skype for Business users from other organizations:'
if ($federationConfig.AllowFederatedUsers) {
  Write-Host -ForegroundColor Green 'The organization allows to communicate with Teams and Skype for Business users from other organizations.'
} else {
  Write-Host -ForegroundColor Red 'The organization does not allow to communicate with Teams and Skype for Business users from other organizations.'
  $toEnableFederation = Read-Host 'Would you like to enable federation? [Y/N]'
  if ($toEnableFederation -eq 'Y') {
    Set-CsTenantFederationConfiguration -AllowFederatedUsers $true
    Write-Host -ForegroundColor Green 'Federation is now enabled.'
  } else {
    return
  }
}

# ALLOW & BLOCK LISTS
$allowedDomains = Get-AllowList
$blockedDomains = Get-BlockList

$isAllowList = Get-IsAllowList $allowedDomains
$isBlockList = Get-IsBlockList $blockedDomains
Write-Host 'Checking if the organization is allowing to communicate with the external domain provided:'
if (!$isAllowList) {
  if (!$isBlockList) {
    Write-Host -ForegroundColor Green 'The organization is allowing communication with all federated domains. (open federation)'
  } else {
    Write-Host 'The organization is blocking communication with the following domains:'
    $parsedBlockList = Get-ParsedBlockList $blockedDomains
    Write-Host $parsedBlockList
    if ($parsedBlockList -contains $Domain) {
      Write-Host -ForegroundColor Red 'The provided domain is in the block list.'
      $toRemoveFromBL = Read-Host 'Would you like to remove it? [Y/N]'
      if ($toRemoveFromBL -eq 'Y') {
        Set-BlockList $parsedBlockList $Domain
      } else { 
        return 
      }
    } else {
      Write-Host -ForegroundColor Green 'The provided domain is not in the block list.'
    }
  }
} else {
  Write-Host -ForegroundColor Yellow 'The organization is not allowing communication with all federated domains. (close federation)'
  $parsedAllowList = Get-ParsedAllowList $allowedDomains
  Write-Host 'The organization is allowing communication only with the following domains:'
  Write-Host $parsedAllowList
  if ($parsedAllowList -contains $Domain) {
    Write-Host -ForegroundColor Green 'The provided domain is included in the allow list.'
  } else {
    Write-Host -ForegroundColor Red 'The provided domain is not included in the allow list.'
    $toAddToAL = Read-Host 'Would you like to add it? [Y/N]'
    if ($toAddToAL -eq 'Y') {
      Set-AllowList $parsedAllowList $Domain
    } else {
      return
    }
  }
  if (!$isBlockList) {
    Write-Host -ForegroundColor Green 'The organization is not blocking communications with any domain.'
  } else {
     Write-Host 'The organization is blocking communication with the following domains:'
     $parsedBlockList = Get-ParsedBlockList $blockedDomains
     Write-Host $parsedBlockList
     if ($parsedBlockList -contains $Domain) {
      Write-Host -ForegroundColor Red 'The provided domain is in the block list.'
      $toRemoveFromBL = Read-Host 'Would you like to remove it? [Y/N]'
      if ($toRemoveFromBL -eq 'Y') {
        Set-BlockList $parsedBlockList $Domain
      } else { 
        return 
      }
     } else {
      Write-Host -ForegroundColor Green 'The provided domain is not in the block list.'
     }
  }
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USER HOMING & COEXISTENCE
$homing = (Get-CsOnlineUser $UPN).HostingProvider

Write-Host 'Checking if the user is homed online in Skype for Business:'
if ($homing -eq 'sipfed.online.lync.com'){
  Write-Host -ForegroundColor Green 'The user is homed online in Skype for Business.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not homed online in Skype for Business user. To federate from Teams, the user has to be moved to the cloud (Move-CsUser).'
}


$coexistence = (Get-CsOnlineUser $UPN).TeamsUpgradeEffectiveMode

Write-Host 'Checking if the user is set up for TeamsOnly:'
if ($coexistence -eq 'TeamsOnly') {
  Write-Host -ForegroundColor Green 'The user is set up for TeamsOnly.'
} else {
  Write-Host -ForegroundColor Red 'The user is not set up for TeamsOnly.'
  $toUpgradeToTeams = Read-Host 'Would you like to upgrade the user to Teams? [Y/N]'
  if ($toUpgradeToTeams -eq 'Y') {
    if ($homing -eq 'sipfed.online.lync.com') {
      Grant-CsTeamsUpgradePolicy -Identity $UPN -PolicyName UpgradeToTeams
      Write-Host -ForegroundColor Green 'The user was upgraded to Teams.'
    } else {
      return Write-Host -ForegroundColor Red 'The user cannot be upgraded to Teams because the user is not homed online in Skype for Business. To be upgraded to Teams, the user has to be moved to the cloud (Move-CsUser).'
    }
  } else {
    return
  }
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# PRIVACY MODE (PRESENCE)
$privacyMode = (Get-CsPrivacyConfiguration).EnablePrivacyMode

Write-Host 'Checking if the tenant is using privacy mode for presence:'
if (!$privacyMode) {
  Write-Host -ForegroundColor Green 'The tenant is not using privacy mode.'
} else {
  Write-Host -ForegroundColor Yellow 'WARNING: The tenant is using privacy mode and external users cannot see tenant users presence.'
  $toDisablePrivacy = Read-Host 'Would you like to disable privacy mode? [Y/N]'
  if ($toDisablePrivacy -eq 'Y') {
    Set-CsPrivacyConfiguration -EnablePrivacyMode $false
    Write-Host -ForegroundColor Green 'The tenant is not using privacy mode anymore.'
  } else {
    return
  }
}

$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  $pName = Get-ExternalAccessPolicyName $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\CannotFederateWithDomain.txt"
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'User:'$UPN
  Write-Host 'External domain:'$Domain
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User $($UPN) (TMS):"
  Get-CsOnlineUser $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User $($UPN) (MSO):"
  Get-MsolUser -UserPrincipalName $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User licenses (MSO):"
  Write-Host "Subscriptions:"$userLicense.SKU
  Write-Host "Service plans:"$userLicense.ServicePlans
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "External access policy (TMS):"
  Get-CsExternalAccessPolicy -Identity $pName
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Tenant federation configuration (TMS):"
  Get-CsTenantFederationConfiguration
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Privacy configuration (TMS):"
  Get-CsPrivacyConfiguration
  Stop-Transcript
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
