$WingetID = "Zoom.Zoom"
$WinGetResolve = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe"
$WinGetPathExe = $WinGetResolve[-1].Path
$WinGetPath = Split-Path -Path $WinGetPathExe -Parent
set-location $WinGetPath
.\winget.exe install -e --id $WingetID -h --accept-package-agreements --accept-source-agreements --scope machine