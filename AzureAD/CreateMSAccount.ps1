<#
Name......: CreateMSAccount.ps1
Version...: 21.08.1
Author....: Dario CORRADA

This script will create a new account on MSOnline and assign Office 365 license
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AzureAD\\CreateMSAccount\.ps1$" > $null
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

$answ = [System.Windows.MessageBox]::Show("Proceed to create Office 365 account?",'ACCOUNT','YesNo','Info')
if ($answ -eq "No") {
    Exit
}

# looking for NuGet Package Provider
try {
    $pp = Get-PackageProvider -Name NuGet
}
catch {
    Install-PackageProvider -Name NuGet -Confirm:$True -MinimumVersion "2.8.5.216" -Force
}

# define username
$form = FormBase -w 520 -h 270 -text "ACCOUNT"
$font = New-Object System.Drawing.Font("Arial", 12)
$form.Font = $font
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(500,30)
$label.Text = "Username:"
$form.Controls.Add($label)
$usrname = New-Object System.Windows.Forms.TextBox
$usrname.Location = New-Object System.Drawing.Point(10,60)
$usrname.Size = New-Object System.Drawing.Size(450,30)
$form.Controls.Add($usrname)
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,100)
$label2.Size = New-Object System.Drawing.Size(500,30)
$label2.Text = "Fullname:"
$form.Controls.Add($label2)
$fullname = New-Object System.Windows.Forms.TextBox
$fullname.Location = New-Object System.Drawing.Point(10,140)
$fullname.Size = New-Object System.Drawing.Size(450,30)
$form.Controls.Add($fullname)
$OKButton = New-Object System.Windows.Forms.Button
OKButton -form $form -x 200 -y 190 -text "Ok"
$form.Topmost = $true
$result = $form.ShowDialog()
$username = $usrname.Text
$completo = $fullname.Text

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
    Get-MsolUser -UserPrincipalName "$username@agmsolutions.net"
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
OKButton -form $form_modalita -x 90 -y 120 -text "Ok"
$result = $form_modalita.ShowDialog()
if ($result -eq "OK") {
    if ($optbasic.Checked) {
        Write-Host -ForegroundColor Blue "Proceed to assign Basic license..."
        $assigned = $label_basic
    } elseif ($consulenti.Checked) {
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
    $form_pswd = FormBase -w 450 -h 230 -text "CREATE PASSWORD"
    $personal = RadioButton -form $form_pswd -checked $true -x 30 -y 20 -text "Set your own password"
    $randomic  = RadioButton -form $form_pswd -checked $false -x 30 -y 50 -text "Generate random password"
    OKButton -form $form_pswd -x 90 -y 120 -text "Ok"
    $result = $form_pswd.ShowDialog()
    if ($result -eq "OK") {
        if ($personal.Checked) {
            $form = FormBase -w 520 -h 200 -text "PASSWORD"
            $font = New-Object System.Drawing.Font("Arial", 12)
            $form.Font = $font
        
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10,20)
            $label.Size = New-Object System.Drawing.Size(500,30)
            $label.Text = "Password:"
            $form.Controls.Add($label)
        
            $usrname = New-Object System.Windows.Forms.TextBox
            $usrname.Location = New-Object System.Drawing.Point(10,60)
            $usrname.Size = New-Object System.Drawing.Size(450,30)
            $usrname.PasswordChar = '*'
            $form.Controls.Add($usrname)
        
            $OKButton = New-Object System.Windows.Forms.Button
        
            OKButton -form $form -x 200 -y 120 -text "Ok"
        
            $form.Topmost = $true
            $result = $form.ShowDialog()

            $thepasswd = $usrname.Text
        } elseif ($randomic.Checked) {
            Add-Type -AssemblyName 'System.Web'
            $thepasswd = [System.Web.Security.Membership]::GeneratePassword(10, 0)
        }
    }
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