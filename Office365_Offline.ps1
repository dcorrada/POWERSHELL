<#
Name......: Office365_Offline.ps1
Version...: 21.2.1
Author....: Dario CORRADA

This script will install Office 365 Business Standard by offline procedure explained as follows

https://support.microsoft.com/en-us/office/use-the-office-offline-installer-f0a85fe7-118f-41cb-a791-d59cef96ad1c?ui=en-us&rs=en-us&ad=us#OfficePlans=signinorgid
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Office365_Offline\.ps1$" > $null
$workdir = $matches[1]

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# temporary working directory
$tmppath = 'C:\ODT'
if (!(Test-Path $tmppath)) {
   New-Item -ItemType directory -Path $tmppath > $null
}

# download and run deployment tool
$download = New-Object net.webclient
$download.Downloadfile("https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_13530-20376.exe", "C:\ODT\Deploy.exe")
[System.Windows.MessageBox]::Show("Target folder to select will be C:\ODT",'INFO','Ok','Info') | Out-Null
Start-Process -FilePath "C:\ODT\Deploy.exe" -Wait
Remove-Item "C:\ODT\Deploy.exe" -Recurse -Force

# create xml config file
'<Configuration>' | Out-File "C:\ODT\installOfficeBusRet64.xml" -Encoding ASCII -Append
'  <Add OfficeClientEdition="64">' | Out-File "C:\ODT\installOfficeBusRet64.xml" -Encoding ASCII -Append
'    <Product ID="O365BusinessRetail">' | Out-File "C:\ODT\installOfficeBusRet64.xml" -Encoding ASCII -Append
'      <Language ID="it-it" />' | Out-File "C:\ODT\installOfficeBusRet64.xml" -Encoding ASCII -Append
'    </Product>' | Out-File "C:\ODT\installOfficeBusRet64.xml" -Encoding ASCII -Append
'  </Add>' | Out-File "C:\ODT\installOfficeBusRet64.xml" -Encoding ASCII -Appends
'</Configuration>' | Out-File "C:\ODT\installOfficeBusRet64.xml" -Encoding ASCII -Append

# download Office
Write-Host -NoNewline "Downloading Office 365 Business..."
$xml_conf = ('/download', 'C:\ODT\installOfficeBusRet64.xml')
Start-Process -Wait "C:\ODT\setup.exe" $xml_conf
Write-Host -ForegroundColor Green " DONE"

# configuring Office
Write-Host -NoNewline "Configuring Office 365 Business..."
$xml_conf = ('/configure', 'C:\ODT\installOfficeBusRet64.xml')
Start-Process -Wait "C:\ODT\setup.exe" $xml_conf
Write-Host -ForegroundColor Green " DONE"

# cleaning temporary
$answ = [System.Windows.MessageBox]::Show("Clean temporary files?",'TEMPUS','YesNo','Info')
if ($answ -eq "Yes") {
    Remove-Item "C:\ODT" -Recurse -Force
} else {
    [System.Windows.MessageBox]::Show("Temporary files are stored in C:\ODT",'INFO','Ok','Info') | Out-Null
}

[System.Windows.MessageBox]::Show("That's all, you have to activate Office",'INFO','Ok','Info') | Out-Null