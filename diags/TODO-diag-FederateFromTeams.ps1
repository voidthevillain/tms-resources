# requires Connect-MicrosoftTeams, Connect-MsolService

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN,
  $Domain
)

$ErrorActionPreference = 'SilentlyContinue' # ??? for line 89

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

$allowedDomains = ($federationConfig.AllowedDomains | Out-String).split("<")[1].split(" ")[0] 
$blockedDomains = ($federationConfig.BlockedDomains | Out-String)

# if ($blockedDomains -eq $null) {
  # Write-Host $federationConfig.BlockedDomains[0].length
  # Write-Host $blockedDomains
  # Write-Host $federationConfig.BlockedDomains
# }

if ($allowedDomains -eq 'AllowAllKnownDomains') {
  if ($federationConfig.BlockedDomains[0].length -eq 0) {
    Write-Host 'The organization allows federation with all domains (open federation).'
  } else {
    Write-Host 'The organization is blocking federation with the following domains:'
    # Write-Host $federationConfig.BlockedDomains
    $domains = $blockedDomains
    Write-Host $domains
  }  
} else {
  Write-Host 'The organization does not allow federation with all domains (closed federation).'
  $allowedDomains = $federationConfig.AllowedDomains | Out-String

  If ($allowedDomains.split(":")[0].trim() -eq 'AllowedDomain') {
    Write-Host "The organization only allows federation with the domains included in the allow list:"
    $domains = $allowedDomains.split(":")[1].split("{")[1].split("}")[0]
    $domainsList = @()
    $count = $domains.split(",").length

    for ($i = 0; $i -le $count; $i++) {
      $domainsList += $domains.split(",")[$i]
    }

    $domainsListX = @()

    foreach ($dmn in $domainsList) {
      $DomainsListX += $dmn.split("=")[1]
    }

    # Write-Host $count, $domains
    Write-Host $domainsListX
  } else {

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

# Set-CsTenantFederationConfiguration -BlockedDomains @{Add=$x}
# Set-CsTenantFederationConfiguration -BlockedDomains @{Add=$y}
# Set-CsTenantFederationConfiguration -BlockedDomains $null


# open federation
# $x = New-CsEdgeAllowAllKnownDomains
# Set-CsTenantFederationConfiguration -AllowedDomainsAsAList $x

