<#
Name......: Check_Credentials.ps1
Version...: 22.12.1
Author....: Dario CORRADA

This script just check if login credentials are correct
#>

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
$workdir -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Locales$" > $null
$repopath = $matches[1]
Import-Module -Name "$repopath\Modules\Forms.psm1"

# get credentials
$form_PWD = FormBase -w 300 -h 230 -text "LOGIN"
Label -form $form_PWD -x 10 -y 20 -text 'Insert your credentials' | Out-Null 
Label -form $form_PWD -x 10 -y 50 -w 80 -text 'Username:' | Out-Null 
$textBox = TxtBox -form $form_PWD -x 90 -y 50 -w 180
Label -form $form_PWD -x 10 -y 80 -w 80 -text 'Password:' | Out-Null 
$MaskedTextBox = TxtBox -form $form_PWD -x 90 -y 80 -w 180 -masked $true
$CheckBox = CheckBox -form $form_PWD -checked $false -x 10 -y 110 -text "Domain Account"
OKButton -form $form_PWD -x 75 -y 150 -text 'Ok' | Out-Null
$result = $form_PWD.ShowDialog()

# get domain name
$thiscomputer = Get-WmiObject -Class Win32_ComputerSystem

if ($CheckBox.Checked) {
    $usr = $thiscomputer.Domain + '\' + $textBox.Text
} else {
    $usr = $textBox.Text
}
$pwd = $MaskedTextBox.Text

[reflection.assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") > $null
$principalContext = [System.DirectoryServices.AccountManagement.PrincipalContext]::new([System.DirectoryServices.AccountManagement.ContextType]'Machine',$env:COMPUTERNAME)


if ($principalContext.ValidateCredentials($usr,$pwd)) {
    Write-Host -ForegroundColor Green "ACCESS GRANTED!"
} else {
    Write-Host -ForegroundColor Red "ACCESS DENIED!"
}

Pause