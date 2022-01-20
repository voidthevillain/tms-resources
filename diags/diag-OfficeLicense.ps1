## Requires Connect-MsolService

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $UPN
)

$isLicensed = (Get-MsolUser -UserPrincipalName $UPN).isLicensed

if ($isLicensed) {
  Write-Host -ForegroundColor Green "`nThe user"$UPN" is licensed." "`n"
} else {
  return Write-Host -ForegroundColor Red  "`nThe user"$UPN" is not licensed."
}

$SKUs = (Get-MsolUser -UserPrincipalName $UPN).Licenses.AccountSkuId #.split(":")[1]
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

Write-Host "Subscriptions:" 
Write-Host $licenses.SKU "`n----------------" 
Write-Host "Service plans:" 
Write-Host $licenses.ServicePlans
