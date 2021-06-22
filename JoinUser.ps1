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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\JoinUser\.ps1$" > $null
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
    # import Active Directory module
    $ErrorActionPreference= 'Stop'
    try {
        Import-Module ActiveDirectory
    } catch {
        Install-Module ActiveDirectory -Confirm:$False -Force
        Import-Module ActiveDirectory
    }
    $ErrorActionPreference= 'Inquire'

    # dialog form
    $form = FormBase -w 520 -h 200 -text "ACCOUNT"
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
    $OKButton = New-Object System.Windows.Forms.Button
    OKButton -form $form -x 200 -y 100 -text "Ok"
    $form.Topmost = $true
    $result = $form.ShowDialog()

    # searching user
    $username = $usrname.Text
    try {
        $usrinfo = Get-ADUser -Identity $username
    }
    catch {
        [System.Windows.MessageBox]::Show("User not found",'ACCOUNT','Ok','Warning')
    }

    # granting local admin privileges
    if ($usrinfo) {
        try {
            Add-LocalGroupMember -Group "Administrators" -Member "AGM\$username"
        }
        catch {
            [System.Windows.MessageBox]::Show("Cannot granting admin privilege to $username",'ACCOUNT','Ok','Error')
        }
    }
}