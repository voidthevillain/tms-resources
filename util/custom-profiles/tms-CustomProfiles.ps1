# credit: https://gist.github.com/DanielSmon/cc3fa072857f0272257a5fd451768c3a

[CmdletBinding()]
Param (
  [Parameter(Mandatory=$true)]
  [String]
  $Name
)

$profileName = $Name
Write-Output "Launching $profileName Teams Profile ..."

$userProfile = $env:USERPROFILE
$appDataPath = $env:LOCALAPPDATA
$customProfile = "$appDataPath\Microsoft\Teams\CustomProfiles\$profileName"
$downloadPath = Join-Path $customProfile "Downloads"

if (!(Test-Path -PathType Container $downloadPath)) {
  New-Item $downloadPath -ItemType Directory |
    Select-Object -ExpandProperty FullName
}

$env:USERPROFILE = $customProfile
Start-Process `
  -FilePath "$appDataPath\Microsoft\Teams\Update.exe" `
  -ArgumentList '--processStart "Teams.exe"' `
  -WorkingDirectory "$appDataPath\Microsoft\Teams"
