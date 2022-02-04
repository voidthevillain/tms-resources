# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Dependencies: WinOS, script from https://github.com/voidthevillain/tms-diags/blob/main/util/custom-profiles/tms-CustomProfiles.ps1
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# USAGE: 
# .\tms-CreateProfile.ps1 profileName
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $Name
)

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# FUNCTIONS
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
function Get-ProfileScript {
  $path = "$($HOME)\Documents\WindowsPowerShell"
  $exists = $false

  gci -path $path | foreach {
    if ($_.Name -eq 'tms-CustomProfiles.ps1') {
      $exists = $true
    } 
  }

  return $exists
}

function New-ProfileScript {
  $path = "$($HOME)\Documents\WindowsPowerShell"
  $name = "tms-CustomProfiles.ps1"

  $url = "https://raw.githubusercontent.com/voidthevillain/tms-diags/main/util/custom-profiles/tms-CustomProfiles.ps1"

  try {
    (curl $url).Content | Out-File -FilePath "$($path)\$($name)"
    return $true
  } catch {
    return $false
  }
}

function Get-CustomProfilesPath {
  $path = "$($env:localappdata)\Microsoft\Teams\CustomProfiles"

  if ($path) {
    return $true
  } else {
    return $false
  }
}

function Get-CustomProfiles {
  $profiles = gci -path "$($env:localappdata)\Microsoft\Teams\CustomProfiles"

  return $profiles
}

function New-CustomProfile {
  param (
    $pName
  )

  $scriptPath = "$($HOME)\Documents\WindowsPowerShell\tms-CustomProfiles.ps1"
  
  PowerShell.exe -File $scriptPath $pName
  }

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
Write-Host 'Checking if profile script exists:'
$profileScript = Get-ProfileScript

if ($profileScript) {
  Write-Host -ForegroundColor Green "Profile script exists in $($HOME)\Documents\WindowsPowerShell."
} else {
  Write-Host 'Profile script does not exist. Downloading script from https://github.com/voidthevillain/tms-diags/blob/main/util/custom-profiles/tms-CustomProfiles.ps1'
  $isCreated = New-ProfileScript
  if ($isCreated) {
    Write-Host -ForegroundColor Green "Successfully downloaded and installed script in $($HOME)\Documents\WindowsPowerShell."
  } else {
    return Write-Host -ForegroundColor Red "Could not download or install script from https://github.com/voidthevillain/tms-diags/blob/main/util/custom-profiles/tms-CustomProfiles.ps1."
  }
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
Write-Host 'Checking if custom profiles exist:'
$profilesExist = Get-CustomProfilesPath
if ($profilesExist) {
  Write-Host 'Custom profiles exist. Getting profiles:'
  $customProfiles = Get-CustomProfiles
  $customProfilesParsed= @()
  foreach ($cp in $customProfiles) {
    Write-Host $cp
    $customProfilesParsed += ($cp.Name | Out-String).trim()
  }

  if ($customProfilesParsed -contains $Name) {
    Write-Host "Profile $($Name) aleady exists. Launching it"
    New-CustomProfile $Name
  } else {
    Write-Host "Profile $($Name) does not exist. Creating and launching it"
    New-CustomProfile $Name
  }
} else {
  Write-Host "Custom profiles do not exist. Creating and launching profile $($Name)"
  New-CustomProfile $Name
}
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
