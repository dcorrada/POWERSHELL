<#
Name......: LockedAccount.ps1
Version...: 19.12.1
Author....: Dario CORRADA

Questo script accede ad Active Directory, cerca gli account bloccati e permette di sbloccarli
#>

$ErrorActionPreference= 'Inquire'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

Import-Module -Name '\\192.168.2.251\Dario\SCRIPT\Moduli_PowerShell\Forms.psm1'

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

# Controllo accesso
$AD_login = LoginWindow

# Importo il modulo di Active Directory
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory } 

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
    $form.Text = "SBLOCCA UTENTE"
    $form.Size = "400,200"
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(350,30)
    $label.Text = "Inserire il nome utente da sbloccare:"
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
            Write-Host -ForegroundColor Green "$username sbloccato"
        }
        Catch { 
            [System.Windows.MessageBox]::Show("Impossibile sbloccare $username",'ATTENZIONE','Ok','Error') > $null
        }
        $ErrorActionPreference = 'Inquire'
    }    
}