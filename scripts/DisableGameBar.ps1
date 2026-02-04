#Disables Game Bar in windows
#Created by ChristopherWStevenson

New-Item -Path HKLM:\Software\Policies\Microsoft\Windows -name GameDVR -force
set-itemproperty -path "HKLM:\Software\Policies\Microsoft\Windows\GameDVR" -name "AllowgameDVR" -value 0