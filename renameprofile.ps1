# close Outlook
Stop-Process -Name "outlook"

# rename e-Mail-profile
$currentProfile = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles").Profiles
Rename-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\$currentProfile" -NewName "old_profile"

# create new e-Mail-Profile
$newProfile = "new_profile"
New-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles" -Name $newProfile

# start Outlook
Start-Process "outlook"

#Please note that this script is intended as an example only and may require adjustments to run successfully in your environment :)
