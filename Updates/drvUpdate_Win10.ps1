<#
Name......: drvUpdate_Win10.ps1
Version...: 22.12.1
Author....: Dario CORRADA

This script tries to fetch and install (optional) driver updates
see also https://learn.microsoft.com/en-us/answers/questions/301795/automate-optional-drivers-install.html
#>

# check execution policy
foreach ($item in (Get-ExecutionPolicy -List)) {
    if(($item.Scope -eq 'LocalMachine') -and ($item.ExecutionPolicy -cne 'Bypass')) {
        Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:`n"
        Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force`n"
        Write-Host -NoNewline "Afterwards restart this script. "
        Pause
        Exit
    }
}

# elevated script execution with admin privileges
$ErrorActionPreference= 'Stop'
try {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if ($testadmin -eq $false) {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
        exit $LASTEXITCODE
    }
}
catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}
$ErrorActionPreference= 'Inquire'

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# looking for NuGet Package Provider
try {
    $pp = Get-PackageProvider -Name NuGet
}
catch {
    Install-PackageProvider -Name NuGet -Confirm:$True -MinimumVersion "2.8.5.216" -Force
}

$ErrorActionPreference= 'Stop'
try {
    Import-Module PSWindowsUpdate
} catch {
    Install-Module PSWindowsUpdate -Confirm:$False -Force
    Import-Module PSWindowsUpdate
}
$ErrorActionPreference= 'Inquire'

# register to MS Update Service
Add-WUServiceManager -ServiceID "7971f918-a847-4430-9279-4a52d1efe18d" -Confirm:$false

# install the drivers updates twice
Install-WindowsUpdate -AcceptAll -Install -IgnoreReboot -Confirm:$False -UpdateType Driver -MicrosoftUpdate -ForceDownload -ForceInstall -ErrorAction SilentlyContinue | Out-File "$env:USERPROFILE\Downloads\$(get-date -f yyyy-MM-dd)-drvupdate.log" -force
Install-WindowsUpdate -AcceptAll -Install -IgnoreReboot -Confirm:$False -UpdateType Driver -MicrosoftUpdate -ForceDownload -ForceInstall -ErrorAction SilentlyContinue | Out-File "$env:USERPROFILE\Downloads\$(get-date -f yyyy-MM-dd)-drvupdate.log" -force

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}
