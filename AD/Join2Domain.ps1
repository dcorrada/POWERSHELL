<#
Name......: Join2Domain.ps1
Version...: 24.09.1
Author....: Dario CORRADA

This script joins a PC to a network domain
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
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

# just pipe more than single "Split-Path" if the script maps to nested subfolders
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent

# graphical stuff
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

<# *******************************************************************************
                                    BODY
******************************************************************************* #>

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
$consulenti  = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "Asset Consulenti"
$milano = RadioButton -form $form_modalita -checked $false -x 30 -y 80 -text "Asset Milano"
$torino = RadioButton -form $form_modalita -checked $false -x 30 -y 110 -text "Asset Torino"
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
        $outarget = 'OU=Consulenti,OU=Computers,OU=Delegate' + $dnsuffix
    } elseif ($milano.Checked) {
        $outarget = 'OU=Milano,OU=Computers,OU=Delegate' + $dnsuffix
    } elseif ($torino.Checked) {
        $outarget = 'OU=Torino,OU=Computers,OU=Delegate' + $dnsuffix
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