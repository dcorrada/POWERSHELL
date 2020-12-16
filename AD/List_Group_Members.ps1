<#
Name......: List_Group_Members.ps1
Version...: 19.10.1
Author....: Dario CORRADA

Questo script serve elenca tutti i membri appartenenti ad un Gruppo di Active Directory

+++ UPDATES +++

[2019-10-03 Dario CORRADA] 
Prima release

#>
$ErrorActionPreference= 'Inquire'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

# Controllo accesso
Import-Module -Name '\\itmilitgroup\SD_Utilities\SCRIPT\Moduli_PowerShell\Patrol.psm1'
$login = Patrol -scriptname List_Group_Members

# Importo il modulo di Active Directory
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory }

# recupero l'elenco dei gruppi in AD
$ADgroups = Get-ADGroup -Filter *
$sorted = $ADgroups.Name | Sort-Object

$form = New-Object System.Windows.Forms.Form
$form.Text = "LISTA GRUPPI"
$form.Size = New-Object System.Drawing.Size(400,230)
$form.StartPosition = 'CenterScreen'

$font = New-Object System.Drawing.Font("Arial", 12)
$form.Font = $font
    
$DropDown = new-object System.Windows.Forms.ComboBox
$DropDown.Location = new-object System.Drawing.Size(10,60)
$DropDown.Size = new-object System.Drawing.Size(350,30)
$DropDown.AutoSize = $true
foreach ($profilo in $sorted) { $DropDown.Items.Add($profilo)  > $null }
$Form.Controls.Add($DropDown)
    
$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
$DropDownLabel.size = new-object System.Drawing.Size(280,30) 
$DropDownLabel.Text = "Selezionare il gruppo"
$Form.Controls.Add($DropDownLabel)
    
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(100,110)
$OKButton.Size = New-Object System.Drawing.Size(75,30)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)
    
$form.Topmost = $true
   
$form.Add_Shown({$DropDown.Select()})
$result = $form.ShowDialog()

$groupname = $DropDown.Text

# recupero la lista dei membri del gruppo
$ADmembers = Get-ADGroupMember -id $groupname -Recursive 


$outfile = "C:\Users\$env:USERNAME\Desktop\ADGroup_members.csv"
"Name;Type;OrganizationalUnit" | Out-File $outfile -Encoding ASCII -Append
foreach ($member in $ADmembers) {
    $member.distinguishedName -match ",OU=([a-zA-Z_\-\.\s0-9]+)," > $null
    $ou = $matches[1]
    
    $new_record = @(
        $member.Name,
        $member.objectClass,
        $ou
    )
    $new_string = [system.String]::Join(";", $new_record)
    $new_string | Out-File $outfile -Encoding ASCII -Append
}