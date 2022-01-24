# Windows

$teams = Get-Process Teams -ErrorAction SilentlyContinue
$outlook = Get-Process Outlook -ErrorAction SilentlyContinue

# check if Outlook is running and quit it to clear add-in cache
if ($outlook) {
  Write-Host 'Outlook is running. Quitting Outlook.'
  $outlook | Stop-Process -Force
  Start-Sleep 2
} else {
  Write-Host 'Outlook is not running.'
}

# check if Teams is running and quit it to clear cache
if ($teams) {
  Write-Host 'Teams is running. Quitting Teams'
  $teams | Stop-Process -Force
  Start-Sleep 2
} else {
  Write-Host 'Teams is not running.'
}

# clear cache
gci -path $env:AppData\Microsoft\Teams | foreach { Remove-Item $_.FullName -Recurse -Force }
Write-Host 'Cache cleared.'

# start Outlook
$toStartOutlook = Read-Host 'Do you wish to start Outlook? [Y/N]'
if ($toStartOutlook -eq 'Y') {
  Start-Process Outlook
}


# start Teams
$toStartTeams = Read-Host 'Do you wish to start Teams? [Y/N]'
if ($toStartTeams -eq 'Y') {
  Start-Process -File "$($env:localappdata)\Microsoft\Teams\Update.exe" -ArgumentList '--processStart "Teams.exe"'
}
