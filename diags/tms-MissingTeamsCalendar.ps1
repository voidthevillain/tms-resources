# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# DEPENDENCIES: Connect-MsolService, Connect-ExchangeOnline
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USAGE: 
# Connect-MsolService
# Connect-ExchangeOnline
# Connect-MicrosoftTeams
# .\tms-TeamsCalendar.ps1 user@domain.com
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN
)

Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
Write-Host 'User:'$UPN
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

function Get-AppSetupPolicyName {
  param (
    [string]$UserPrincipalName
  )

  $pN = (Get-CsOnlineUser $UserPrincipalName).TeamsAppSetupPolicy

   if ($pN -eq $null) {
    $pN = 'Global'
  }

  return $pN
}

function Get-IsAppInUserAppSetupPolicy {
  param (
    [string]$UserPrincipalName,
    [string]$AppId
  )

  $policyName = (Get-CsOnlineUser $UserPrincipalName).TeamsAppSetupPolicy

  if ($policyName -eq $null) {
    $policyName = 'Global'
  }

  $policyApps = (Get-CsTeamsAppSetupPolicy -Identity $policyName).PinnedAppBarApps

  foreach ($app in $policyApps) {
    if ($app.Id -eq $calendarId) {
      return $true
    }
  }

  return $false
}
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# User & Licenses
$user = Get-MsolUser -UserPrincipalName $UPN

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

Write-Host 'Checking user licenses:'

# check Teams (TEAMS1)
if ($userLicense.ServicePlans -contains 'TEAMS1') {
  Write-Host -ForegroundColor Green 'The user is licensed for Teams.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Teams.'
}

# check SfBO (MCOSTANDARD, MCO_TEAMS_IW)
if ($userLicense.ServicePlans -contains 'MCOSTANDARD' -OR $userLicense.ServicePlans -contains 'MCO_TEAMS_IW' ) {
  Write-Host -ForegroundColor Green 'The user is licensed for Skype for Business Online.'
} else {
  # Unsure whether to return - in case the user is homed in SfB server
  # return Write-Host 'The user is not licensed for Skype for Business Online.'
  Write-Host -ForegroundColor Red 'The user is not licensed for Skype for Business Online.'
}

# check Exchange Online (EXCHANGE_S_STANDARD, EXCHANGE_S_ENTERPRISE) 
if ($userLicense.ServicePlans -contains 'EXCHANGE_S_STANDARD' -OR $userLicense.ServicePlans -contains 'EXCHANGE_S_ENTERPRISE') {
  Write-Host -ForegroundColor Green 'The user is licensed for Exchange Online.'
} else {
  return Write-Host -ForegroundColor Red 'The user is not licensed for Exchange Online.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Coexistence mode
$coexistenceMode = (Get-CsOnlineUser $UPN).TeamsUpgradeEffectiveMode

Write-Host 'Checking the coexistence mode of the user:'
if ($coexistenceMode -eq 'SfBOnly' -OR $coexistenceMode -eq 'SfBWithTeamsCollab') {
  return Write-Host -ForegroundColor Red 'The coexistence mode assigned to the user does not support Calendar in Teams:' $coexistenceMode
} else {
  Write-Host -ForegroundColor Green 'The coexistence mode assigned to the user supports Calendar in Teams:' $coexistenceMode
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# App setup policy
$calendarId = 'ef56c0de-36fc-4ef8-b417-3d82ba9d073c'
$appInPolicy = Get-IsAppInUserAppSetupPolicy $UPN $calendarId
Write-Host 'Checking if the app setup policy of the user includes Calendar:'

if ($appInPolicy) {
  Write-Host -ForegroundColor Green 'The Calendar app is included in the app setup policy of the user.'
} else {
  return Write-Host -ForegroundColor Red 'The Calendar app is not included in the app setup policy of the user.'
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Exchange homing
$recipientType = (Get-Mailbox -Identity $UPN).RecipientTypeDetails
Write-Host 'Checking the Exchange homing of the user:'

if ($recipientType -eq 'UserMailbox') {
  Write-Host -ForegroundColor Green 'The mailbox is hosted in Exchange Online.'
} elseif ($recipientType -eq 'MailUser') {
  return Write-Host -ForegroundColor Yellow 'The mailbox is hosted in Exchange On-Premises. See https://docs.microsoft.com/en-us/microsoftteams/troubleshoot/exchange-integration/teams-exchange-interaction-issue'
}


# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# EWS settings for tenant and user
# check tenant and user EWS settings and offer corrections
$tenantEwsEnabled = (Get-OrganizationConfig).EwsEnabled
Write-Host 'Checking tenant EWS settings:'

if ($tenantEwsEnabled -eq $null -OR $tenantEwsEnabled -eq $true) {
  Write-Host -ForegroundColor Green 'Tenant EWS is enabled.'
} else {
  $toEnableTenantEWS = Read-Host 'Tenant EWS is disabled. Would you like to enable it? [Y/N]'
  if ($toEnableTenantEWS -eq 'Y') {
    Set-OrganizationConfig -EwsEnabled $true 
    Write-Host -ForegroundColor Green 'Tenant EWS is now enabled.'
  } else {
    return
  }
}

# tenant EWS restrictions
$tenantEwsApplicationAccessPolicy = (Get-OrganizationConfig).EwsApplicationAccessPolicy

if ($tenantEwsApplicationAccessPolicy -eq $null) {
  Write-Host -ForegroundColor Green 'The tenant does not restrict EWS access.'
} elseif ($tenantEwsApplicationAccessPolicy -eq 'EnforceAllowList') {
  $toRemoveTenantEWSAllowList = Read-Host 'The tenant is allowing EWS access only for the applications in the allow list. Would you like to remove the restriction? [Y/N]'
  if ($toRemoveTenantEWSAllowList -eq 'Y') {
    Set-OrganizationConfig -EwsApplicationAccessPolicy $null
    Write-Host -ForegroundColor Green 'The tenant does not restrict EWS access anymore.'
  } else {
    return
  }
} elseif ($tenantEwsApplicationAccessPolicy -eq 'EnforceBlockList') {
  $toRemoveTenantEWSBlockList = Read-Host 'The tenant is blocking EWS access for the applications in the block list. Would you like to remove the restriction? [Y/N]'
  if ($toRemoveTenantEWSBlockList -eq 'Y') {
    Set-OrganizationConfig -EwsApplicationAccessPolicy $null
    Write-Host -ForegroundColor Green 'The tenant does not restrict EWS access anymore.'
  } else {
    return
  }
}

# user EWS
$userEwsEnabled = (Get-CASMailbox -Identity $UPN).EwsEnabled
Write-Host 'Checking user EWS settings:'

if ($userEwsEnabled -eq $null -OR $userEwsEnabled -eq $true) {
  Write-Host -ForegroundColor Green 'Mailbox EWS is enabled.'
} else {
  $toEnableUserEWS = Read-Host 'Mailbox EWS is disabled. Would you like to enable it? [Y/N]'
  if ($toEnableUserEWS -eq 'Y') {
    Set-CASMailbox -Identity $UPN -EwsEnabled $true
    Write-Host -ForegroundColor Green 'Mailbox EWS is now enabled.'
  } else {
    return
  }
}

# user EWS restrictions
$userEwsApplicationAccessPolicy = (Get-CASMailbox -Identity $UPN).EwsApplicationAccessPolicy

if ($userEwsApplicationAccessPolicy -eq $null) {
  Write-Host -ForegroundColor Green 'The mailbox does not restrict EWS access.'
} elseif ($userEwsApplicationAccessPolicy -eq 'EnforceAllowList') {
  $toRemoveMbxEWSAllowList = Read-Host 'The mailbox is allowing EWS access only for the applications in the allow list. Would you like to remove the restriction? [Y/N]'
  if ($toRemoveMbxEWSAllowList -eq 'Y') {
    Set-CASMailbox -Identity $UPN -EwsApplicationAccessPolicy $null
    Write-Host -ForegroundColor Green 'The mailbox does not restrict EWS access anymore.'
  } else {
    return
  }
} elseif ($userEwsApplicationAccessPolicy -eq 'EnforceBlockList') {
  $toRemoveMbxEWSBlockList = Read-Host 'The mailbox is blocking EWS access for the applications in the block list. Would you like to remove the restriction? [Y/N]'
  if ($toRemoveMbxEWSBlockList -eq 'Y') {
    Set-CASMailbox -Identity $UPN -EwsApplicationAccessPolicy $null
    Write-Host -ForegroundColor Green 'The mailbox does not restrict EWS access anymore.'
  } else {
    return
  }
}

$issuePersists = Read-Host 'No issues found. Does the problem persist? [Y/N]'
if ($issuePersists -eq 'Y') {
  $desktopPath = [Environment]::GetFolderPath("Desktop")
  $pName = Get-AppSetupPolicyName $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Generating transcript:'
  Start-Transcript -Path "$($desktopPath)\MissingTeamsCalendar.txt"
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
  Write-Host "App setup policy apps (TMS):"
  (Get-CsTeamsAppSetupPolicy -Identity $pName).PinnedAppBarApps
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "User $($UPN) mailbox (EXO):"
  Get-Mailbox -Identity $UPN
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Tenant EWS settings (EXO):"
  Get-OrganizationConfig | select *ews*
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host "Mailbox EWS settings (EXO):"
  Get-CASMailbox -Identity $UPN | select *ews*
  Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
  Write-Host 'Please open a support request with the above transcript attached.'
} else {
  return
}
Write-Host '...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -.'
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
