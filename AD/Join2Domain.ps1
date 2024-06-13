<#
Name......: Join2Domain.ps1
Version...: 20.12.1
Author....: Dario CORRADA

This script joins a PC to a network domain
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AD\\Join2Domain\.ps1$" > $null
$workdir = $matches[1]

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

$hostname = $env:computername

# getting domain name
$output = nslookup ls
$output[0] -match "Server:\s+[a-zA-Z_\-0-9]+\.([a-zA-Z\-0-9\.]+)$" > $null
$dominio = $matches[1]
$answ = [System.Windows.MessageBox]::Show("Join to [$dominio]?",'DOMAIN','YesNo','Info')
if ($answ -eq "No") {    
    $form = FormBase -w 300 -h 175 -text "DOMAIN"
    Label -form $form -x 10 -y 20 -text 'Domain name:' | Out-Null
    $adom = TxtBox -form $form -x 10 -y 50 -w 250
    OKButton -form $form -x 100 -y 90 -text "Ok" | Out-Null
    $result = $form.ShowDialog()
    $dominio = $adom.Text 
}

# getting AD credentials
# starting release 24.05.1 credentials are managed from PSWallet
Write-Host -NoNewline "Credential management... "
$pswout = PowerShell.exe -file "$workdir\Safety\Stargate.ps1" -ascript 'Join2Domain'
if ($pswout.Count -eq 2) {
    $ad_login = New-Object System.Management.Automation.PSCredential($pswout[0], (ConvertTo-SecureString $pswout[1] -AsPlainText -Force))
} else {
    [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "Ko"
    Pause
    exit
}
Write-Host -ForegroundColor Green 'Ok'

# OU dialog box
$form_modalita = FormBase -w 300 -h 230 -text "OU DESTINATION"
$noou = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "null"
$consulenti  = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "Client Consulenti"
$milano = RadioButton -form $form_modalita -checked $false -x 30 -y 80 -text "Client Milano"
$torino = RadioButton -form $form_modalita -checked $false -x 30 -y 110 -text "Client Torino"
OKButton -form $form_modalita -x 90 -y 150 -text "Ok" | Out-Null
$result = $form_modalita.ShowDialog()

# get distinguished name suffix
$dnsuffix = ''
foreach ($dctag in $dominio.Split('.')) {
    $dnsuffix += ',DC=' + $dctag
}

# in "elseif" blocks modify $outarget prefix according to yours OU paths
if ($result -eq "OK") {
    if ($noou.Checked) {
        $outarget = "null"
    } elseif ($consulenti.Checked) {
        $outarget = 'OU=Client Consulenti,OU=Delegate' + $dnsuffix
    } elseif ($milano.Checked) {
        $outarget = 'OU=Client Milano,OU=Delegate' + $dnsuffix
    } elseif ($torino.Checked) {
        $outarget = 'OU=Client Torino,OU=Delegate' + $dnsuffix
    }    
}

Write-Host "Domain...: " -NoNewline
Write-Host $dominio -ForegroundColor Cyan
Write-Host "OU.......: " -NoNewline
Write-Host $outarget -ForegroundColor Cyan

$ErrorActionPreference= 'Stop'
Try {
    if ($outarget -eq "null") {
        Add-Computer -ComputerName $hostname -Credential $ad_login -DomainName $dominio -Force
    } else {
        Add-Computer -ComputerName $hostname -Credential $ad_login -DomainName $dominio -OUPath $outarget -Force
    }
    Write-Host "PC joined to domain" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
} 

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}