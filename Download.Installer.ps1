$WingetChrome = .\winget.exe show -e --id Google.Chrome
$url = ( -split ($WingetChrome | Select-String "Installer Url*"))[-1]
$version = ( -split ($WingetChrome | Select-String "Version*"))[-1]
$dir = "C:\ProgramData\BCIT-Cache"
$file = "$($dir)\Chrome.v$($version)-Enterprise.msi"
$webClient = New-Object System.Net.WebClient
if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force }
$webClient.DownloadFile($url, $file)