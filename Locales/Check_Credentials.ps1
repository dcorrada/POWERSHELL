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
$form_PWD = FormBase -w 400 -h 270 -text "LOGIN"
$label = Label -form $form_PWD -x 10 -y 20 -w 300 -h 20 -text 'Insert your credentials'
$usrlabel = Label -form $form_PWD -x 10 -y 50 -w 100 -h 20 -text 'Username:'
$textBox = TxtBox -form $form_PWD -x 130 -y 50 -w 150 -h 20 -text ''
$CheckBox = CheckBox -form $form_PWD -checked $false -x 20 -y 120 -text "Domain Account"
$pwdlabel = Label -form $form_PWD -x 10 -y 80 -w 100 -h 20 -text 'Password:'
$MaskedTextBox = New-Object System.Windows.Forms.MaskedTextBox
$MaskedTextBox.PasswordChar = '*'
$MaskedTextBox.Location = New-Object System.Drawing.Point(130,80)
$MaskedTextBox.Size = New-Object System.Drawing.Size(150,20)
$form_PWD.Add_Shown({$MaskedTextBox.Select()})
$form_PWD.Controls.Add($MaskedTextBox)
$OKButton = OKButton -form $form_PWD -x 100 -y 160 -text 'Ok'
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