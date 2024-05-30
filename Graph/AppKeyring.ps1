param (
    [Parameter(Mandatory=$true)]
    [string]$ascript
)

<#
Name......: AppKeyring.ps1
Version...: 24.06.1
Author....: Dario CORRADA

This script is a frontend for external PSWallet calls, dedicated for the 
credential management of regitered Azure apps
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Graph\\AppKeyring\.ps1$" > $null
$workdir = $matches[1]

# header 
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

$WhereIsMyWallet =  "$workdir\Safety\PSWallet.ps1"
$WhereIsMyDB = $env:LOCALAPPDATA + '\PSWallet.sqlite'
$idi = @{
    UPN    = 'null'
    CLIENT = 'null'
    TENANT = 'null'
    SECRET = 'null'
}
if ((Test-Path -Path $WhereIsMyWallet -PathType Leaf) -and (Test-Path -Path $WhereIsMyDB -PathType Leaf)) {
    [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Open Key File"
    $OpenFileDialog.initialDirectory = "$env:LOCALAPPDATA"
    $OpenFileDialog.filter = 'Key file (*.key)| *.key'
    $OpenFileDialog.ShowDialog() | Out-Null
    $WhereIsMyKey = $OpenFileDialog.filename 

    $AccessForm = FormBase -w 280 -h 180 -text 'ACCESS'
    $addnewentry = RadioButton -form $AccessForm -x 20 -y 10 -checked $false -text 'Add new credential'
    $listentries = RadioButton -form $AccessForm -x 20 -y 40 -checked $true -text 'Select existing credential'
    OKButton -form $AccessForm -x 75 -y 95 -text "Next" | Out-Null
    $resultButton = $AccessForm.ShowDialog()
    if ($addnewentry.Checked) {
        $pswout = PowerShell.exe -file $WhereIsMyWallet `
                -ExtKey $WhereIsMyKey  `
                -ExtScript $ascript  `
                -ExtAction 'addGraph'
        
        $pswout | foreach {
            if ($_ -match "^PSWallet>>> \[APPID\] (.+)$") {
                $idi.CLIENT = "$($matches[1])"
            } elseif ($_ -match "^PSWallet>>> \[TENANTID\] (.+)$") {
                $idi.TENANT = "$($matches[1])"
            } elseif ($_ -match "^PSWallet>>> \[SECRETVALUE\] (.+)$") {
                $idi.SECRET = "$($matches[1])"
            } elseif ($_ -match "^PSWallet>>> \[UPN\] (.+)$") {
                $idi.UPN = "$($matches[1])"
            } else {
                $matches = @()
            }
        }
    } elseif ($listentries.Checked) {
        $pswout = PowerShell.exe -file $WhereIsMyWallet `
                -ExtKey $WhereIsMyKey  `
                -ExtScript $ascript  `
                -ExtAction 'listGraph'

        <# 
        Fare un form con due bottoni: uno per accedere, l'altro per aggiornare 
        la secret scaduta (mostrare data di scadenza sul form)
        #>
    }
} else {
    [System.Windows.MessageBox]::Show("No PSWallet or related database found",'ABORTING','Ok','Error') | Out-Null
}

Write-Host -ForegroundColor Blue "$($idi.UPN)"
Write-Host -ForegroundColor Blue "$($idi.CLIENT)"
Write-Host -ForegroundColor Blue "$($idi.TENANT)"
Write-Host -ForegroundColor Blue "$($idi.SECRET)"