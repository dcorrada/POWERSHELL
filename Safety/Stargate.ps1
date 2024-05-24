param (
    [Parameter(Mandatory=$true)]
    [string]$ascript
)

<#
Name......: Stargate.ps1
Version...: 24.05.1
Author....: Dario CORRADA

This script is a minimal frontend for external PSWallet calls
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Safety\\Stargate\.ps1$" > $null
$workdir = $matches[1]

# header 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

$AccessForm = FormBase -w 300 -h 210 -text 'ACCESS'
Label -form $AccessForm -x 10 -y 20 -w 80 -text 'Username:' | Out-Null
$usrname = TxtBox -form $AccessForm -x 100 -y 20 -w 170
Label -form $AccessForm -x 10 -y 50 -w 80 -text 'Password:' | Out-Null
$passwd = TxtBox -form $AccessForm -x 100 -y 50 -w 170 -masked $true
$pswadd = CheckBox -form $AccessForm -x 10 -y 80 -text 'Add new credential to PSWallet'
RETRYButton -form $AccessForm -x 150 -y 120 -w 120 -text "PSWallet user list" | Out-Null
OKButton -form $AccessForm -x 20 -y 120 -w 120 -text "Directly access" | Out-Null
$resultButton = $AccessForm.ShowDialog()

$WhereIsMyWallet =  "$workdir\Safety\PSWallet.ps1"
if ($resultButton -eq 'RETRY') {
    if (Test-Path -Path $WhereIsMyWallet -PathType Leaf) {
        [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Title = "Open Key File"
        $OpenFileDialog.initialDirectory = "$env:LOCALAPPDATA"
        $OpenFileDialog.filter = 'Key file (*.key)| *.key'
        $OpenFileDialog.ShowDialog() | Out-Null
        $WhereIsMyKey = $OpenFileDialog.filename 

        $pswout = PowerShell.exe -file $WhereIsMyWallet `
            -ExtKey $WhereIsMyKey  `
            -ExtScript $ascript  `
            -ExtAction 'listusr'
        $userlist = @()
        foreach ($currentItem in $pswout) {
            if ($currentItem -match "PSWallet>>> (.+)$") {
                $currentItem -match "PSWallet>>> (.+)$" | Out-Null
                $userlist += $matches[1]
            }
        }

        $Formlist = FormBase -w 275 -h ((($userlist.Count) * 30) + 120) -text "SELECT"
        $they = 20
        $choices = @()
        foreach ($remote in $userlist) {
            if ($they -eq 20) {
                $isfirst = $true
            } else {
                $isfirst = $false
            }
            $choices += RadioButton -form $Formlist -x 25 -y $they -checked $isfirst -text $remote
            $they += 25 
        }
        OKButton -form $Formlist -x 75 -y ($they + 30) -text "Ok" | Out-Null
        $result = $Formlist.ShowDialog()
        foreach ($item in $choices) {
            if ($item.Checked) {
                $ausr = $item.Text
            }
        }

        if ($ausr -eq 'NO DATA FOUND') {
            exit
        } else {
            $pswout = PowerShell.exe -file $WhereIsMyWallet `
                -ExtKey $WhereIsMyKey  `
                -ExtScript $ascript  `
                -ExtUsr $ausr  `
                -ExtAction 'getpwd'
            foreach ($currentItem in $pswout) {
                if ($currentItem -match "PSWallet>>> (.+)$") {
                    $currentItem -match "PSWallet>>> (.+)$" | Out-Null
                    $apwd = $matches[1]
                }
            }
        }
    }
} else {
    $ausr = $usrname.Text
    $apwd = $passwd.Text
    if (($pswadd.Checked) -and (Test-Path -Path $WhereIsMyWallet -PathType Leaf)) {
        [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Title = "Open Key File"
        $OpenFileDialog.initialDirectory = "$env:LOCALAPPDATA"
        $OpenFileDialog.filter = 'Key file (*.key)| *.key'
        $OpenFileDialog.ShowDialog() | Out-Null
        $WhereIsMyKey = $OpenFileDialog.filename 

        $pswout = PowerShell.exe -file $WhereIsMyWallet `
                -ExtKey $WhereIsMyKey  `
                -ExtScript $ascript  `
                -ExtUsr $ausr  `
                -ExtPwd $passwd.Text  `
                -ExtAction 'add'
    }
}

Write-Host "$ausr"
Write-Host "$apwd"