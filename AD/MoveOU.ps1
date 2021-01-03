<#
Name......: SpostamentoOU.ps1
Version...: 19.04.1
Author....: Dario CORRADA

This script read a computer list from a file, and move these computer from a OU to another
#>


# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
$workdir -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AD$" > $null
$repopath = $matches[1]
Import-Module -Name "$repopath\Modules\Forms.psm1"


# get AD credentials
$AD_login = LoginWindow

# Import Active Directory module
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory } 

# retrieve computer list
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms')
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Text file (*.txt)| *.txt'
$OpenFileDialog.ShowDialog() | Out-Null
$file_path = $OpenFileDialog.filename
$computer_list = Get-Content $file_path

# retrieve available OUs
$ou_available = Get-ADOrganizationalUnit -Filter *
$ou_list = @()
foreach ($item in $ou_available) {
    $ou_list += $item.DistinguishedName
}

$source_dest = @()
foreach ($item in ('OU SOURCE', 'OU DESTINATION')) {
    $formlist = FormBase -w 400 -h 200 -text $item
    $DropDown = new-object System.Windows.Forms.ComboBox
    $DropDown.Location = new-object System.Drawing.Size(10,60)
    $DropDown.Size = new-object System.Drawing.Size(350,30)
    foreach ($elem in ($ou_list | sort)) {
        $DropDown.Items.Add($elem)  > $null
    }
    $formlist.Controls.Add($DropDown)
    $DropDownLabel = new-object System.Windows.Forms.Label
    $DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
    $DropDownLabel.size = new-object System.Drawing.Size(500,30) 
    $DropDownLabel.Text = "Select OU"
    $formlist.Controls.Add($DropDownLabel)
    OKButton -form $formlist -x 100 -y 100 -text "Ok"
    $formlist.Add_Shown({$DropDown.Select()})
    $result = $formlist.ShowDialog()
    $source_dest += $DropDown.Text
}

foreach ($computer_name in $computer_list) {
    Write-Host -Nonewline $computer_name
    $computer_ADobj = Get-ADComputer $computer_name -Credential $AD_login
    # Write-Host $computer_ADobj.DistinguishedName
    if ($computer_ADobj.DistinguishedName -match $source_dest[1]) {
        Write-Host -ForegroundColor Cyan " skipped"
    } elseif ($computer_ADobj.DistinguishedName -match $source_dest[0]) {
        $target_path = "OU=" + $dest_ou + $suffix
        $computer_ADobj | Move-ADObject -Credential $AD_login -TargetPath $source_dest[1]
        Write-Host -ForegroundColor Green " remapped"
    } else {
        Write-Host -ForegroundColor Cyan " skipped"
    }
}
pause