# Wipe any prior Winget PSObject
Remove-Variable * -ErrorAction SilentlyContinue
# Enter Winget ID
$WingetID = "Zoom.Zoom"
# Gather Winget Information
$WinGetResolve = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe"
$WinGetPathExe = $WinGetResolve[-1].Path
$WinGetPath = Split-Path -Path $WinGetPathExe -Parent
set-location $WinGetPath
$WingetInfo = .\winget.exe show -e --id $WingetID --scope machine
$WingetInfoUser = .\winget.exe show -e --id $WingetID --scope User
$WingetInstalled = .\winget.exe list -e --id $WingetID --source winget
#Check ID
if ($WingetInfo[-1] -ne "No package found matching input criteria.") {
    
    # Workstation Information
    $Update = (-split $WingetInstalled[-3])[-1]
    if ($Update -ne "Available") { $Update = "No Update Available" }
    elseif ((-split $WingetInstalled)[-2] -ne $WingetID) { $Installed = $false }
    else { $Installed = $true }
    if ($Installed = $true) { $InstalledVersion = (-split $WingetInstalled)[-1] }
    else { $InstalledVersion = $null }

    # Create Winget Properties
    $WingetApp = if (($WingetInfo | Select-String "Found *")) { ((($WingetInfo | Select-String "Found *")[0].ToString()).Trim("Found")).Trim("[$($WingetID)]").Trim() }
    $WingetData = New-Object PSObject ; $WingetInfo | Where-Object { $_ -match ': ' } | ForEach-Object { $Item = $_.Trim() -split ':\s'; $WingetData | Add-Member -MemberType NoteProperty -Name $($Item[0] -replace '[:\s]', '') -Value $Item[1] -EA SilentlyContinue }
    
    if (($WingetInfo | Select-String "Release Notes: *") -and ($WingetInfo | Select-String "Purchase Url: *")) {
        $ReleaseNotes = "(?<=Release Notes:)(.*?)(?=Purchase Url:)"
        $SelectedText = [regex]::Matches($WingetInfo, $ReleaseNotes)
        $ReleaseNotes = $SelectedText[0].Value.Trim()
    }
    if (($WingetInfo | Select-String "Documentation: *") -and ($WingetInfo | Select-String "Tags: *")) {
        $Documentation = "(?<=Documentation:)(.*?)(?=Tags:)"
        $SelectedText = [regex]::Matches($WingetInfo, $Documentation)
        $Documentation = $SelectedText[0].Value.Trim()
    }
    if (($WingetInfo | Select-String "Tags: *") -and ($WingetInfo | Select-String "Installer: *")) {
    $Tags = "(?<=Tags:)(.*?)(?=Installer:)"
    $SelectedText = [regex]::Matches($WingetInfo, $Tags)
    $Tags = $SelectedText[0].Value.Trim()
    $Tags = (-split $Tags)
    }
    $MachineScope = ( -split ($WingetInfo | Select-String "Installer Type:*"))[-1]
    $InstallerTypeUser = ( -split ($WingetInfoUser | Select-String "Installer Type:*"))[-1]
    $InstallerURLUSer = ( -split ($WingetInfoUser | Select-String "Installer URL:*"))[-1]
    $InstallerSHA256User = ( -split ($WingetInfoUser | Select-String "Installer SHA256:*"))[-1]
    if ($MachineScope) { $MachineScope = $true } else { $MachineScope = $false }
    if ($InstallerTypeUser) { $UserScope = $true } else { $UserScope = $false }

    # Generate PSObject
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("Application", "$WingetApp"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("WingetID", "$WingetID"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("Installed", "$Installed"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("InstalledVersion", "$InstalledVersion"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("Update", "$Update"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("ReleaseNotes", "$ReleaseNotes"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("Documentation", "$Documentation"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("Tags", "$Tags"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("MachineScope", "$MachineScope"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("UserScope", "$UserScope"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("UserInstallerType", "$InstallerTypeUser"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("UserInstallerURL", "$InstallerURLUser"))
    $WingetData.psobject.Properties.Add([PSNoteProperty]::new("UserInstallerSHA256", "$InstallerSHA256User"))
    
    $WingetData | Format-List

}
else {
    Write-host $WingetInfo[-1] -ForegroundColor Red
}