<#
Name......: AddMailAccount.ps1
Version...: 23.1.1
Author....: Dario CORRADA

This script will manually configure a IMAP/SMTP account onto Windos Mail or MS Outlook 
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\miscellaneous\\AddMailAccount\.ps1$" > $null
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

# check if Outlook Desktop installed
$MailClient = Get-ItemProperty HKLM:\Software\Classes\mailto\shell\open\command | Select-Object -ExpandProperty '(default)'
if (!$MailClient) { 
    $MailClient = Get-ItemProperty HKLM:\Software\Classes\mailto\shell\open\command | Select-Object -ExpandProperty '(default)' 
}

# getting account info
$aform = FormBase -w 300 -h 300 -text "MAIL SERVERS"
Label -form $aform -x 10 -y 20 -w 100 -h 30 -text 'IMAP server:' | Out-Null
$iname = TxtBox -form $aform -x 110 -y 20 -w 150 -h 30 -text 'imap.foobar.baz'
Label -form $aform -x 10 -y 50 -w 100 -h 30 -text 'IMAP port:' | Out-Null
$iport = TxtBox -form $aform -x 110 -y 50 -w 150 -h 30 -text '993'
Label -form $aform -x 10 -y 80 -w 100 -h 30 -text 'IMAP security:' | Out-Null
$idrop = DropDown -form $aform -x 110 -y 80  -w 150 -h 30 -opts ('SSL/TLS', 'STARTTLS')
Label -form $aform -x 10 -y 110 -w 100 -h 30 -text 'SMTP server:' | Out-Null
$sname = TxtBox -form $aform -x 110 -y 110 -w 150 -h 30 -text 'smtp.foobar.baz'
Label -form $aform -x 10 -y 140 -w 100 -h 30 -text 'SMTP port:' | Out-Null
$sport = TxtBox -form $aform -x 110 -y 140 -w 150 -h 30 -text '465'
Label -form $aform -x 10 -y 170 -w 100 -h 30 -text 'SMTP security:' | Out-Null
$sdrop = DropDown -form $aform -x 110 -y 170  -w 150 -h 30 -opts ('SSL/TLS', 'STARTTLS')
$sdrop.Text = 'STARTTLS'
OKButton -form $aform -x 75 -y 220 -text "Ok"  | Out-Null
$result = $aform.ShowDialog()
$srvinfo = @{
    ImapName = $iname.Text
    ImapPort = $iport.Text
    ImapSec = $idrop.Text
    SmtpName = $sname.Text
    SmtpPort = $sport.Text
    SmtpSec = $sdrop.Text
}
$usrlogin = LoginWindow


# configuring specific client
<#
Sembra che non ci sia modo di configurare account IMAP/SMTP via riga di comando
Al limite si può passare dal tool del pannello di controllo il cui eseguibile 
per Outlook si trova come C:\Program Files\Microsoft Office\root\Office16\OLCFG.EXE
#>
if ($MailClient -match "Office16\\OUTLOOK\.EXE$") {
    Write-Host -NoNewline "Configuring Microsoft Outlook 365... "
} else {
    [System.Windows.MessageBox]::Show("No compliant client found",'ABORT','Ok','Warning') | Out-Null
}

<#
Managing autodiscover.xml 
* https://exchangepedia.com/2015/10/use-a-powershell-function-to-get-autodiscover-xml.html
* https://4sysops.com/archives/control-outlook-autodiscover-using-registry-and-powershell/
#>

# close Outlook client first...
$OutlookApplication = New-Object -comobject Outlook.Application # close Outlook client first
$autodiscoverxml = $outlookApplication.Session.AutoDiscoverXML

