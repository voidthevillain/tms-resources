## Requires Connect-MsolService, Connect-ExchangeOnline

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN
)

####################### LICENSE #########################
# check if licensed, won't proceed if unlicensed
$isLicensed = (Get-MsolUser -UserPrincipalName $UPN).isLicensed

if ($isLicensed) {
  Write-Host 'The user is licensed.'
} else {
  return Write-Host 'The user is not licensed.'
}
#########################################################
####################### HOMING #########################
# check homing, won't proceed if homed on-premises
$recipientType = (Get-Mailbox -Identity $UPN).RecipientTypeDetails

if ($recipientType -eq 'UserMailbox') {
  Write-Host 'The mailbox is hosted in Exchange Online.'
} elseif ($recipientType -eq 'MailUser') {
  return Write-Host 'The mailbox is hosted in Exchange On-Premises. See https://docs.microsoft.com/en-us/microsoftteams/troubleshoot/exchange-integration/teams-exchange-interaction-issue'
}
#########################################################
####################### SERVICE #########################
# check if has service plans: TMS, SfBO, EXO
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

# check Teams (TEAMS1)
if ($licenses.ServicePlans -contains 'TEAMS1') {
  Write-Host 'The user is licensed for Teams.'
} else {
  return Write-Host 'The user is not licensed for Teams.'
}

# check SfBO (MCOSTANDARD, MCO_TEAMS_IW)
if ($licenses.ServicePlans -contains 'MCOSTANDARD' -OR $licenses.ServicePlans -contains 'MCO_TEAMS_IW' ) {
  Write-Host 'The user is licensed for Skype for Business Online.'
} else {
  # Unsure whether to return - in case the user is homed in SfB server
  # return Write-Host 'The user is not licensed for Skype for Business Online.'
  Write-Host 'The user is not licensed for Skype for Business Online.'
}

# check Exchange Online (EXCHANGE_S_STANDARD, EXCHANGE_S_ENTERPRISE) 
if ($licenses.ServicePlans -contains 'EXCHANGE_S_STANDARD' -OR $licenses.ServicePlans -contains 'EXCHANGE_S_ENTERPRISE') {
  Write-Host 'The user is licensed for Exchange Online.'
} else {
  return Write-Host 'The user is not licensed for Exchange Online.'
}

#########################################################
####################### EWS T&U #########################
# check tenant and user EWS settings and offer corrections
$tenantEwsEnabled = (Get-OrganizationConfig).EwsEnabled

if ($tenantEwsEnabled -eq $null -OR $tenantEwsEnabled -eq $true) {
  Write-Host 'Tenant EWS is enabled.'
} else {
  $toEnableTenantEWS = Read-Host 'Tenant EWS is disabled. Would you like to enable it? [Y/N]'
  if ($toEnableTenantEWS -eq 'Y') {
    Set-OrganizationConfig -EwsEnabled $true 
    Write-Host 'Tenant EWS is now enabled.'
  } else {
    return
  }
}

$tenantEwsApplicationAccessPolicy = (Get-OrganizationConfig).EwsApplicationAccessPolicy

if ($tenantEwsApplicationAccessPolicy -eq $null) {
  Write-Host 'The tenant does not restrict EWS access.'
} elseif ($tenantEwsApplicationAccessPolicy -eq 'EnforceAllowList') {
  $toRemoveTenantEWSAllowList = Read-Host 'The tenant is allowing EWS access only for the applications in the allow list. Would you like to remove the restriction? [Y/N]'
  if ($toRemoveTenantEWSAllowList -eq 'Y') {
    Set-OrganizationConfig -EwsApplicationAccessPolicy $null
    Write-Host 'The tenant does not restrict EWS access anymore.'
  } else {
    return
  }
} elseif ($tenantEwsApplicationAccessPolicy -eq 'EnforceBlockList') {
  $toRemoveTenantEWSBlockList = Read-Host 'The tenant is blocking EWS access for the applications in the block list. Would you like to remove the restriction? [Y/N]'
  if ($toRemoveTenantEWSBlockList -eq 'Y') {
    Set-OrganizationConfig -EwsApplicationAccessPolicy $null
    Write-Host 'The tenant does not restrict EWS access anymore.'
  } else {
    return
  }
}

$userEwsEnabled = (Get-CASMailbox -Identity $UPN).EwsEnabled

if ($userEwsEnabled -eq $null -OR $userEwsEnabled -eq $true) {
  Write-Host 'Mailbox EWS is enabled.'
} else {
  $toEnableUserEWS = Read-Host 'Mailbox EWS is disabled. Would you like to enable it? [Y/N]'
  if ($toEnableUserEWS -eq 'Y') {
    Set-CASMailbox -Identity $UPN -EwsEnabled $true
    Write-Host 'Mailbox EWS is now enabled.'
  } else {
    return
  }
}

$userEwsApplicationAccessPolicy = (Get-CASMailbox -Identity $UPN).EwsApplicationAccessPolicy

if ($userEwsApplicationAccessPolicy -eq $null) {
  Write-Host 'The mailbox does not restrict EWS access.'
} elseif ($userEwsApplicationAccessPolicy -eq 'EnforceAllowList') {
  $toRemoveMbxEWSAllowList = Read-Host 'The mailbox is allowing EWS access only for the applications in the allow list. Would you like to remove the restriction? [Y/N]'
  if ($toRemoveMbxEWSAllowList -eq 'Y') {
    Set-CASMailbox -Identity $UPN -EwsApplicationAccessPolicy $null
    Write-Host 'The mailbox does not restrict EWS access anymore.'
  } else {
    return
  }
} elseif ($userEwsApplicationAccessPolicy -eq 'EnforceBlockList') {
  $toRemoveMbxEWSBlockList = Read-Host 'The mailbox is blocking EWS access for the applications in the block list. Would you like to remove the restriction? [Y/N]'
  if ($toRemoveMbxEWSBlockList -eq 'Y') {
    Set-CASMailbox -Identity $UPN -EwsApplicationAccessPolicy $null
    Write-Host 'The mailbox does not restrict EWS access anymore.'
  } else {
    return
  }
}
#########################################################
