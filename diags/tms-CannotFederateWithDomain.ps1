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
# TODO: check homing, coexistence, and external access policy 
# MAYBE: remove from block list and add to allow list???

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN,
  $Domain
)

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
    $licenses.SKU += $SKUs.split(":")[1]
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
# FEDERATION
$federationConfig = Get-TenantFederationConfig

Write-Host 'Checking if the organization allows to communicate with Teams and Skype for Business users from other organizations:'
if ($federationConfig.AllowFederatedUsers) {
  Write-Host -ForegroundColor Green 'The organization allows to communicate with Teams and Skype for Business users from other organizations.'
} else {
  return Write-Host -ForegroundColor Red 'The organization does not allow to communicate with Teams and Skype for Business users from other organizations.'
  # TODO: prompt to enable federation ?
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
      return Write-Host -ForegroundColor Red 'The provided domain is in the block list.'
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
    return Write-Host -ForegroundColor Red 'The provided domain is not included in the allow list.'
  }

  if (!$isBlockList) {
    Write-Host -ForegroundColor Green 'The organization is not blocking communications with any domain.'
  } else {
     Write-Host 'The organization is blocking communication with the following domains:'
     $parsedBlockList = Get-ParsedBlockList $blockedDomains
     Write-Host $parsedBlockList
     if ($parsedBlockList -contains $Domain) {
      return Write-Host -ForegroundColor Red 'The provided domain is in the block list.'
     } else {
      Write-Host -ForegroundColor Green 'The provided domain is not in the block list.'
     }
  }
}

Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'



# allow list
# $list = New-Object Collections.Generic.List[String]
# $list.add("contoso.com")
# $list.add("fabrikam.com")
# $list.add("zapp.com")
# $list.add("pyru.com")
# $list.add("pyruz.com")
# $list.add("pyruzz.com")
# Set-CsTenantFederationConfiguration -AllowedDomainsAsAList $list

# block list
# $x = New-CsEdgeDomainPattern -Domain "contoso.com"
# $y = New-CsEdgeDomainPattern -Domain "fabrikam.com"
# $z = New-CsEdgeDomainPattern -Domain "pogg.com"

# Set-CsTenantFederationConfiguration -BlockedDomains @{Add=$x}
# Set-CsTenantFederationConfiguration -BlockedDomains @{Add=$y}
# Set-CsTenantFederationConfiguration -BlockedDomains @{Add=$z}

# empty block list
# Set-CsTenantFederationConfiguration -BlockedDomains $null


# open federation
# $x = New-CsEdgeAllowAllKnownDomains
# Set-CsTenantFederationConfiguration -AllowedDomainsAsAList $x


# any mess?
# enable federation
# Set-CsTenantFederationConfiguration -AllowFederatedUsers $true
# empty block list
# Set-CsTenantFederationConfiguration -BlockedDomains $null
# empty allow list
# $x = New-CsEdgeAllowAllKnownDomains
# Set-CsTenantFederationConfiguration -AllowedDomainsAsAList $x
