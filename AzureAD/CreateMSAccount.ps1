<#
Name......: CreateMSAccount.ps1
Version...: 21.08.1
Author....: Dario CORRADA

This script will create a new account on MSOnline and assign Office 365 license
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
# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

$answ = [System.Windows.MessageBox]::Show("Proceed to create Office 365 account?",'ACCOUNT','YesNo','Info')
if ($answ -eq "No") {
    Exit
}

# define username
$form = FormBase -w 350 -h 270 -text "ACCOUNT"
Label -form $form -x 20 -y 20 -w 80 -text 'Username:' | Out-Null
$usrname = TxtBox -form $form -x 100 -y 20 
Label -form $form -x 20 -y 50 -w 80 -text 'Fullname:' | Out-Null
$fullname = TxtBox -form $form -x 100 -y 50 
$personal = RadioButton -form $form -checked $true -x 20 -y 80 -text "Set your own password"
$apass = TxtBox -form $form -x 40 -y 110 -w 260 -masked $true
$randomic  = RadioButton -form $form -checked $false -x 20 -y 140 -text "Generate random password"
OKButton -form $form -x 120 -y 190 -text "Ok" | Out-Null
$result = $form.ShowDialog()
$username = $usrname.Text
$completo = $fullname.Text
if ($personal.Checked) {
    $thepasswd = $apass.Text
} elseif ($randomic.Checked) {
    Add-Type -AssemblyName 'System.Web'
    $thepasswd = [System.Web.Security.Membership]::GeneratePassword(10, 0)
}

# import the AzureAD module
$ErrorActionPreference= 'Stop'
try {
    Import-Module MSOnline
} catch {
    Install-Module MSOnline -Confirm:$False -Force
    Import-Module MSOnline
}
$ErrorActionPreference= 'Inquire'

# connect to Tenant
Connect-MsolService

# check for existing account
$ErrorActionPreference= 'Stop'
try {
    # re-edit your current domain suffix before running this script
    Get-MsolUser -UserPrincipalName "$username@foobarbaz.net"
    [System.Windows.MessageBox]::Show("$username already exists! Aborting...",'WARNING','Ok','Warning')
    Exit
}
catch {
    Write-Host -ForegroundColor Blue "Proceed to create $username..."
}
$ErrorActionPreference= 'Inquire'

# check for available licenses
$avail_licenses = Get-MsolAccountSku
foreach ($license in $avail_licenses) {
    if ($license.AccountSkuId -match "O365_BUSINESS_ESSENTIALS$") {
        $label_basic = $license.AccountSkuId
        $avail_basic = $license.ActiveUnits - $license.ConsumedUnits
    } elseif ($license.AccountSkuId -match "O365_BUSINESS_PREMIUM$") {
        $avail_standard = $license.ActiveUnits - $license.ConsumedUnits
        $label_standard = $license.AccountSkuId
    }
}

# select license to assign
$form_modalita = FormBase -w 300 -h 230 -text "OFFICE 365 LICENSE"
$optbasic = RadioButton -form $form_modalita -checked $false -x 30 -y 20 -text "Office 365 Basic"
if ($avail_basic -lt 1) { $optbasic.Enabled = $false }
$optstandard  = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "Office 365 Standard"
if ($avail_standard -lt 1) { $optstandard.Enabled = $false }
OKButton -form $form_modalita -x 90 -y 120 -text "Ok" | Out-Null
$result = $form_modalita.ShowDialog()
if ($result -eq "OK") {
    if ($optbasic.Checked) {
        Write-Host -ForegroundColor Blue "Proceed to assign Basic license..."
        $assigned = $label_basic
    } elseif ($optstandard.Checked) {
        Write-Host -ForegroundColor Blue "Proceed to assign Standard license..."
        $assigned = $label_standard
    } else {
        [System.Windows.MessageBox]::Show("No license assigned! Aborting...",'WARNING','Ok','Warning')
        Exit
    }
}

# create account
($firstname,$lastname) = $completo.Split(' ')
$ErrorActionPreference= 'Stop'
try {
    $domains = Get-MsolDomain
    $suffix = $domains[($domains.Count - 1)].Name
    New-MsolUser -UserPrincipalName "$username@$suffix" -FirstName $firstname -LastName $lastname -DisplayName $completo -Password $thepasswd -ForceChangePassword $false -UsageLocation 'IT' -LicenseAssignment $assigned
}
catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}
$ErrorActionPreference= 'Inquire'

Write-Host -ForegroundColor Green "ACCOUNT CREATED!"
Write-Host -ForegroundColor Cyan "$completo <$username@$suffix>"
Write-Host -ForegroundColor Blue -NoNewline "Password: "
Write-Host "$thepasswd"
Write-Host -ForegroundColor Blue -NoNewline "License: "
Write-Host "$assigned"
Pause