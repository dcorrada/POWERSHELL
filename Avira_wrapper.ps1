<#
Name......: Avira_wrapper.ps1
Version...: 22.05.1
Author....: Dario CORRADA

This script automatically fetch and install Avira host, launche the UI and then unistalls it
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

Write-Host -ForegroundColor Blue "*** Avira ***`n"

# Download and install
Write-Host -NoNewline 'Download and install... '
$download = New-Object net.webclient
$downbin = 'C:\Users\' + $env:USERNAME + '\Downloads\avira.exe'
$download.DownloadFile('https://package.avira.com/download/connect-client-win/package/avira_it_swu_1897812318-1649253704__pswuws.exe', $downbin)
#Invoke-WebRequest -Uri 'https://package.avira.com/download/connect-client-win/package/avira_it_swu_1897812318-1649253704__pswuws.exe' -OutFile $downbin
Start-Process -Wait -FilePath $downbin '/S'
Remove-Item $downbin -Force 
Write-Host -ForegroundColor Green "DONE`n"

# Finding binary
$apath = Get-ChildItem -Path 'C:\Program Files' -Filter 'Avira.Systray.exe' -Force -Recurse -ErrorAction SilentlyContinue
if ($apath) {
    $fullpath = $apath.DirectoryName
} else {
    $apath = Get-ChildItem -Path 'C:\Program Files (x86)' -Filter 'Avira.Systray.exe' -Force -Recurse -ErrorAction SilentlyContinue
    if ($apath) {
        $fullpath = $apath.DirectoryName
    } else {
        [System.Windows.MessageBox]::Show("Ccleaner not found",'ERROR','Ok','Error') > $null
        Exit
    }
}

# Ccleaner launch
Write-Host -NoNewline 'Perform... '
$thebin = $fullpath + '\Avira.Systray.exe'
Start-Process $thebin '/showMiniGui'
Write-Host -ForegroundColor Green "DONE`n"

# Uninstall
[System.Windows.MessageBox]::Show("Click Ok to uninstall Avira",'UNINSTALL','Ok','Warning') > $null
Write-Host -NoNewline 'Uninstall... '
$ErrorActionPreference= 'SilentlyContinue'
$prey = Get-Process Avira.SoftwareUpdater.ServiceHost
Stop-Process -Id $prey.Id -Force
Start-Sleep 2
$searchkey = 'Avira Software Updater'
$record = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
          Get-ItemProperty | Where-Object {$_.DisplayName -match $searchkey } | Select-Object -Property DisplayName, UninstallString
foreach ($elem in $record) {
    $elem.UninstallString -match "MsiExec\.exe .+\{([A-Z0-9\-]+)\}" > $null
    $unstring = $Matches[1]
}
msiexec "/X{$unstring}"
$prey = Get-Process Avira.Systray
Stop-Process -Id $prey.Id -Force
Start-Sleep 2
$prey = Get-Process Avira.ServiceHost
Stop-Process -Id $prey.Id -Force
Start-Sleep 2
$apath = Get-ChildItem -Path 'C:\ProgramData' -Filter 'Avira.OE.Setup.Bundle.exe' -Force -Recurse -ErrorAction SilentlyContinue
if ($apath) {
    $fullpath = $apath.DirectoryName
    $thebin = $fullpath + '\Avira.OE.Setup.Bundle.exe'
    Start-Process -Wait $thebin '/uninstall'
}
Write-Host -ForegroundColor Green "DONE`n"
