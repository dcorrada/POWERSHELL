<#
Name......: LockedAccount.ps1
Version...: 19.12.1
Author....: Dario CORRADA

This script looks for locked account and ask to unlock
#>

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
$workdir -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AD$" > $null
$repopath = $matches[1]
Import-Module -Name "$repopath\Modules\Forms.psm1"

# check Active Directory module
if ((Get-Module -Name ActiveDirectory -ListAvailable) -eq $null) {
    $ErrorActionPreference= 'Stop'
    try {
        Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability –Online
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to install RSAT"
        Pause
        Exit
    }
    $ErrorActionPreference= 'Inquire'
}

# get AD credentials
$AD_login = LoginWindow

$locked_accounts = Search-ADAccount -LockedOut

foreach ($elem in $locked_accounts) {
    $OUs = $elem.DistinguishedName.Split(',')
    $locktime = Get-ADUser -Identity $elem.Name -Properties AccountLockoutTime
    Write-Host -NoNewline $locktime.AccountLockoutTime, " "
    foreach ($string in $OUs) {
        $found = $string -match "OU=(.+)"
        if ($found -eq  $true) {
            Write-Host -ForegroundColor Yellow -NoNewline $Matches[1] " "
        }
    }
    Write-Host -ForegroundColor Red $elem.Name
}

Write-Host " "
$result = "OK"
while ($result -eq "OK") {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "ULOCK ACCOUNT"
    $form.Size = "400,200"
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(350,30)
    $label.Text = "Username:"
    $form.Controls.Add($label)
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,60)
    $textBox.Size = New-Object System.Drawing.Size(350,30)
    $form.Controls.Add($textBox)
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = "200,100"
    $CancelButton.Size = '100,30'
    $CancelButton.Text = "Cancel"
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.AcceptButton = $CancelButton
    $form.Controls.Add($CancelButton)
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = "50,100"
    $OKButton.Size = '100,30'
    $OKButton.Text = "Ok"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)
    $form.Add_Shown({$textBox.Select()})
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $username = $textBox.Text
        $ErrorActionPreference = 'Stop'
        Try {
            Unlock-ADAccount –Identity $username -Credential $AD_login
            Write-Host -ForegroundColor Green "$username unlocked"
        }
        Catch { 
            [System.Windows.MessageBox]::Show("Unable to unlock $username",'ERROR','Ok','Error') > $null
        }
        $ErrorActionPreference = 'Inquire'
    }    
}