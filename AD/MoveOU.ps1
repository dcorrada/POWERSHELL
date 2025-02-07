<#
Name......: SpostamentoOU.ps1
Version...: 19.04.1
Author....: Dario CORRADA

This script read a computer list from a file, and move these computer from a OU to another
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

    $formlist = FormBase -w 400 -h 175 -text $item
    Label -form $formlist -x 10 -y 20 -text 'Select OU:'
    $ous = DropDown -form $formlist -x 10 -y 50 -w 350 -opts ($ou_list | sort)
    OKButton -form $formlist -x 140 -y 90 -text "Ok" | Out-Null
    $result = $formlist.ShowDialog()
    $source_dest += $ous.Text
    
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