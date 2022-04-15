<#
Name......: SafetyScan.ps1
Version...: 22.04.1
Author....: Dario CORRADA

This script automatically fetch and performs a scan with Microsoft Safety Scanner Tool
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

Write-Host -ForegroundColor Blue "*** MS Safety Scanner ***`n"

# Download
Write-Host -NoNewline 'Download... '
$download = New-Object net.webclient
$downbin = 'C:\Users\' + $env:USERNAME + '\Downloads\MSERT.exe'
if (Test-Path $downbin) {
    Remove-Item $downbin -Force    
}
$download.DownloadFile('https://go.microsoft.com/fwlink/?LinkId=212732', $downbin)
#Invoke-WebRequest -Uri 'https://go.microsoft.com/fwlink/?LinkId=212732' -OutFile $downbin
Write-Host -ForegroundColor Green "DONE`n"

<#
SYNOPSIS:
/Q      quiet mode, no UI is shown
/N      detect-only mode
/F      force full scan
/F:Y    force full scan and automatic clean
/H      detect high and severe threats only
#>
# perform a quick scan
Write-Host -NoNewline 'Perform... '
Start-Process -Wait $downbin '/N /Q'
Write-Host -ForegroundColor Green "DONE`n"

# log check
notepad 'C:\Windows\debug\msert.log'

# perform a full scan
$answ = [System.Windows.MessageBox]::Show("Do you need a full scan?",'FULL-SCAN','YesNo','Info')
if ($answ -eq "Yes") {    
    Write-Host -NoNewline 'Perform... '
    Start-Process -Wait $downbin '/F:Y /Q'
    Write-Host -ForegroundColor Green "DONE`n"

    # log check
    notepad 'C:\Windows\debug\msert.log'
}


















