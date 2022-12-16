<#
Name......: JoinUser.ps1
Version...: 21.06.1
Author....: Dario CORRADA

This script grants local admin privileges to an existing domain user
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AD\\JoinUser\.ps1$" > $null
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

$answ = [System.Windows.MessageBox]::Show("Are you looking for an AD user?",'ACCOUNT','YesNo','Info')
if ($answ -eq "Yes") {

    # dialog form
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
    $label2.Text = "Password:"
    $form.Controls.Add($label2)
    $passwd = New-Object System.Windows.Forms.MaskedTextBox
    $passwd.PasswordChar = '*'
    $passwd.Location = New-Object System.Drawing.Point(10,140)
    $passwd.Size = New-Object System.Drawing.Size(450,30)
    $form.Controls.Add($passwd)
    $OKButton = New-Object System.Windows.Forms.Button
    OKButton -form $form -x 200 -y 190 -text "Ok"
    $form.Topmost = $true
    $result = $form.ShowDialog()

    # add domain prefix to username
    $username = $usrname.Text
    $thiscomputer = Get-WmiObject -Class Win32_ComputerSystem
    $fullname = $thiscomputer.Domain + '\' + $username

    # test user
    [reflection.assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") > $null
    $principalContext = [System.DirectoryServices.AccountManagement.PrincipalContext]::new([System.DirectoryServices.AccountManagement.ContextType]'Machine',$env:COMPUTERNAME)
    if ($principalContext.ValidateCredentials($fullname,$passwd.Text)) {
        Write-Host -ForegroundColor Green "User OK"
        
        # granting local admin privileges
        try {
            Add-LocalGroupMember -Group "Administrators" -Member $fullname
        }
        catch {
            [System.Windows.MessageBox]::Show("Cannot granting admin privilege to $username",'ACCOUNT','Ok','Error')
        }
    } else {
        [System.Windows.MessageBox]::Show("Invalid credentials for $username",'ACCOUNT','Ok','Error')
    }
}