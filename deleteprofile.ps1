Stop Outlook process
Stop-Process -Name outlook

#Delete current email profile
Remove-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles\YourProfileName" -Recurse

#Create new email profile
$newProfile = New-Item -Path "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows Messaging Subsystem\Profiles" -Name "YourNewProfileName"
Set-ItemProperty -Path $newProfile.PSPath -Name "ProfileName" -Value "YourNewProfileName"

Start Outlook
Start-Process -FilePath "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"

#In the lines above, you need to replace "YourProfileName" with the name of your current email profile and "YourNewProfileName" with the name of your new email profile. Also make sure the path to Outlook is correct in the last line.
