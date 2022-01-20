# requires Connect-MicrosoftTeams, Connect-MsolService
# INCOMPLETE
# TODO: to review function Add-DomainToAllowList, to check homing and coexistence mode
# +++++ TO FURTHER TEST, especially Add-DomainToAllowList for ... appended on 4th item (??? ALWAYS ???)

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN,
  $Domain
)

# $ErrorActionPreference = 'SilentlyContinue' # ??? for line 89

function Output-AllowedDomains {
  param (
    [string[]]$AllowedDomains
  )

  $ads = $AllowedDomains | Out-String

  Try {
    $ads = $ads.split("<")[1].split(" ")[0]
  } Catch {}

  return $ads.trim()
}

function Parse-AllowedDomains {
  param (
    [string[]]$AllowedDomains
  )

  $domains = $allowedDomains.split(":")[1].split("{")[1].split("}")[0]
  $domainsList = @()
  $count = $domains.split(",").length

  for ($i = 0; $i -le $count; $i++) {
    $domainsList += $domains.split(",")[$i]
  }

  $domainsListX = @()

  try {
    foreach ($dmn in $domainsList) {
     $DomainsListX += $dmn.split("=")[1]
    }
  } catch {}

  return $domainsListX
}

function Parse-BlockedDomains {
  param (
    [string[]]$BlockedDomains
  )

  $domains = $BlockedDomains
  $domains = $domains.trimstart().trim()
  $arr = $domains.trim().split("`r")

  $newArr = @()
  for ($i = 0; $i -lt $arr.length; $i += 2) {
    $newArr += $arr[$i]
  }

  $finalArr = @()
  foreach ($item in $newArr) {
    $finalArr += $item.split(":")[1].trimstart()
  }
  
  return $finalArr
}

function Remove-DomainFromBlockList {
    param (
    [string[]]$D
  )

  Set-CsTenantFederationConfiguration -BlockedDomains @{Remove=$D}
}

function Add-DomainToAllowList {
    param (
    $DList,
    $DomainX
  )

  $list = New-Object Collections.Generic.List[String]

  foreach ($dmn in $DList) {
    $list.add($dmn)
  }

  $list.add($DomainX)
  $list[3] = $list[3].trim('...') # item at index 3 gets appended ... ?????????

  # Write-host $list
  # Write-host $list.gettype()
  # write-host $list[3]

  # foreach ($item in $list) {
  #   $index = $list.indexOf($item)
  #   list[$index] = $item.trim('...')
  # }

 
  Set-CsTenantFederationConfiguration -AllowedDomainsAsAList $list
}


$user = (Get-CsOnlineUser $UPN)

if ($user) {
  Write-Host 'The user exists.'
} else {
  return Write-Host 'The user does not exist.'
}

$isLicensed = (Get-MsolUser -UserPrincipalName $UPN).isLicensed

if ($isLicensed) {
  Write-Host 'The user is licensed.'
} else {
  return Write-Host 'The user is not licensed.'
}

$SKUs = (Get-MsolUser -UserPrincipalName $UPN).Licenses.AccountSkuId
$ServicePlans = (Get-MsolUser -UserPrincipalName $UPN).Licenses.ServiceStatus.ServicePlan.ServiceName

$licenses = @{
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

if ($licenses.ServicePlans -contains 'TEAMS1') {
  Write-Host 'The user is licensed for Teams.'
} else {
  return Write-Host 'The user is not licensed for Teams.'
}

if ($licenses.ServicePlans -contains 'MCOSTANDARD' -OR $licenses.ServicePlans -contains 'MCO_TEAMS_IW' ) {
  Write-Host 'The user is licensed for Skype for Business Online.'
} else {
  return Write-Host 'The user is not licensed for Skype for Business Online.'
}

$sipEnabled = $user.Enabled 

if ($sipEnabled) {
  Write-Host 'The user is SIP enabled.'
} else {
  return Write-Host 'The user is not SIP enabld.'
}

$MCOValidationError = $user.MCOValidationError

if (!$MCOValidationError) {
  Write-Host 'The user has no MCOValidationError.'
} else {
  return Write-Host 'The user has an MCOValidationError:' $MCOValidationError
}

$federationConfig = Get-CsTenantFederationConfiguration

if ($federationConfig.AllowFederatedUsers -eq $true) {
  Write-Host 'The organization allows to communicate with Teams and Skype for Business users from other organizations.'
} else {
  $toEnableFederation = Read-Host 'The organization does not allow to communicate with Teams and Skype for Business users from other organizations. Would you like to enable federation? [Y/N]'
  if ($toEnableFederation -eq 'Y') {
    Set-CsTenantFederationConfiguration -AllowFederatedUsers $true
    Write-Host 'The organization now allows to communicate with Teams and Skype for Business users from other organizations.'
  } else {
    return
  }
}

# $allowedDomains = ($federationConfig.AllowedDomains | Out-String).split("<")[1].split(" ")[0] 
$blockedDomains = ($federationConfig.BlockedDomains | Out-String)
$allowedDomains = Output-AllowedDomains $federationConfig.AllowedDomains
# $blockedDomains = Output-BlockedDomains $federationConfig.BlockedDomains

# Write-Host $blockedDomains.Domains
# Write-Host $blockedDomains.LengthOfFirst
# Write-Host $allowedDomains.length

# Write-Host $blockedDomains

# $bds = Output-BlockedDomains $federationConfig.BlockedDomains
# Write-Host 'bds' $bds.Domains


if ($allowedDomains -eq 'AllowAllKnownDomains') {
  if ($federationConfig.BlockedDomains[0].length -eq 0) {
    Write-Host 'The organization allows federation with all domains (open federation).'
  } else {
    Write-Host 'The organization is blocking federation with the following domains:'

        $parsedBlockedDomains = Parse-BlockedDomains $blockedDomains

        Write-Host $parsedBlockedDomains

        if ($parsedBlockedDomains -contains $Domain) {
          $toRemoveDomainFromBlockList = Read-Host 'The organization is blocking federation with the provided domain. Would you like to remove the domain from the block list? [Y/N]'
          if ($toRemoveDomainFromBlockList -eq 'Y') {
            Remove-DomainFromBlockList $Domain
            Write-Host 'The domain was removed from the block list.'
          } else {
            return 'The user cannot federate with externals on the provided domain as it is blocked by the organization.'
          }
        } else {
          Write-Host 'The organization is not blocking federation with the provided domain.'
        }
  }  
} else {
  Write-Host 'The organization does not allow federation with all domains (closed federation).'
  $allowedDomains = $federationConfig.AllowedDomains | Out-String

  If ($allowedDomains.split(":")[0].trim() -eq 'AllowedDomain') {
    Write-Host "The organization only allows federation with the domains included in the allow list:"
    
    $parsedAllowedDomains = Parse-AllowedDomains $allowedDomains

    Write-Host $parsedAllowedDomains

    if ($parsedAllowedDomains -contains $Domain) {
      write-host 'The provided domain is included in the allow list.'
    } else {
      $toAddDomainToAllowList = Read-Host 'The provided domain is not included in the allow list. Would you like to add it? [Y/N]'
      if ($toAddDomainToAllowList -eq 'Y') {
        Add-DomainToAllowList $parsedAllowedDomains $Domain
        Write-Host 'The provided domain is now included in the allow list.'
      } else {
        return 'The user cannot federate with externals on the provided domain as it is not included in the list of domains the organization allows to federate with.'
      }
    }
  } 
}

# allow list
# $list = New-Object Collections.Generic.List[String]
# $list.add("contoso.com")
# $list.add("fabrikam.com")
# $list.add("zapp.com")
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

