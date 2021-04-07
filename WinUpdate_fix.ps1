<#
Name......: WinUpdate_fix.ps1
Version...: 21.04.1
Author....: Dario CORRADA

This script tries to fix corrupted Windows updates
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# stopping services
net stop wuauserv /y
net stop bits /y
net stop appidsvc /y
net stop cryptsvc /y

# backupping cache
$dest_path = 'C:\BACKUP_WindowsUpdate_' + (Get-Date -Format "yyyy.MM.dd-HH.mm")
New-Item -ItemType directory -Path $dest_path
Move-Item -Path "$env:SystemRoot\SoftwareDistribution" -Destination $dest_path -Force
Move-Item -Path "$env:SystemRoot\system32\catroot2" -Destination $dest_path -Force
New-Item -ItemType directory -Path "$env:SystemRoot\SoftwareDistribution"
New-Item -ItemType directory -Path "$env:SystemRoot\system32\catroot2"

# restarting services
net start wuauserv
net start bits
net start appidsvc
net start cryptsvc

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot now?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {
    Restart-Computer
}