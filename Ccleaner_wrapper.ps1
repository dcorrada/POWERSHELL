<#
Name......: Ccleaner_wrapper.ps1
Version...: 22.04.1
Author....: Dario CORRADA

This script automatically fetch and install Ccleaner, performs a scan job and then unistalls it
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

Write-Host -ForegroundColor Blue "*** CCleaner ***`n"

# Download and install
Write-Host -NoNewline 'Download and install... '
$download = New-Object net.webclient
$downbin = 'C:\Users\' + $env:USERNAME + '\Downloads\ccinstall.exe'
$download.DownloadFile('https://bits.avcdn.net/productfamily_CCLEANER/insttype_FREE/platform_WIN_PIR/installertype_ONLINE/build_RELEASE', $downbin)
#Invoke-WebRequest -Uri 'https://bits.avcdn.net/productfamily_CCLEANER/insttype_FREE/platform_WIN_PIR/installertype_ONLINE/build_RELEASE' -OutFile $downbin
Start-Process -Wait -FilePath $downbin '/S'
Remove-Item $downbin -Force 
Write-Host -ForegroundColor Green "DONE`n"

# Finding binary
$apath = Get-ChildItem -Path 'C:\Program Files' -Filter 'Ccleaner.exe' -Force -Recurse -ErrorAction SilentlyContinue
if ($apath) {
    $fullpath = $apath.DirectoryName
} else {
    $apath = Get-ChildItem -Path 'C:\Program Files (x86)' -Filter 'Ccleaner.exe' -Force -Recurse -ErrorAction SilentlyContinue
    if ($apath) {
        $fullpath = $apath.DirectoryName
    } else {
        [System.Windows.MessageBox]::Show("Ccleaner not found",'ERROR','Ok','Error') > $null
        Exit
    }
}

# Ccleaner launch
Write-Host -NoNewline 'Perform... '
$thebin = $fullpath + '\Ccleaner.exe'
Start-Process -Wait $thebin '/AUTO'
$answ = [System.Windows.MessageBox]::Show("Perform registry cleaning?",'REGCLEAN','YesNo','Info')
if ($answ -eq "Yes") {    
    Start-Process $thebin '/REGISTRY'
    [System.Windows.MessageBox]::Show("Click Ok to uninstall Ccleaner",'UNINSTALL','Ok','Warning') > $null
}
Write-Host -ForegroundColor Green "DONE`n"

# Uninstall
Write-Host -NoNewline 'Uninstall... '
$ErrorActionPreference= 'SilentlyContinue'
$prey = Get-Process CCleaner64
Stop-Process -Id $prey.Id -Force
Start-Sleep 2
$ErrorActionPreference= 'Inquire'
$unbin = $fullpath + '\uninst.exe'
Start-Process -Wait $unbin '/S'
Write-Host -ForegroundColor Green "DONE`n"
