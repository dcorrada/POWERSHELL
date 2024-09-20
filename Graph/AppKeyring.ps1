param ([string]$ascript='AppKeyring')

<#
Name......: AppKeyring.ps1
Version...: 24.07.1
Author....: Dario CORRADA

This script is a frontend for external PSWallet calls, dedicated for the 
credential management of regitered Azure apps
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# check execution policy
foreach ($item in (Get-ExecutionPolicy -List)) {
    if(($item.Scope -eq 'LocalMachine') -and ($item.ExecutionPolicy -cne 'Bypass')) {
        Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:`n"
        Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force`n"
        Write-Host -NoNewline "Afterwards restart this script. "
        Pause
        Exit
    }
}

# elevated script execution with admin privileges
$ErrorActionPreference= 'Stop'
try {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if ($testadmin -eq $false) {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
        exit $LASTEXITCODE
    }
}
catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}
$ErrorActionPreference= 'Inquire'

# just pipe more than single "Split-Path" if the script maps to nested subfolders
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent

# graphical stuff
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"


<# *******************************************************************************
                                    BODY
******************************************************************************* #>

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

        $PSWrecords = $pswout | ForEach-Object {
            if ($_ -match "^PSWallet>>> ([a-zA-Z_\-\.0-9@]+) <-> ([a-zA-Z_\-\.0-9]+) ([a-zA-Z\-\s0-9\(\)]+)$") {
                New-Object -TypeName PSObject -Property @{
                    UPN = $matches[1]
                    APP = $matches[2]
                    EXP = $matches[3]
                } | Select UPN, APP, EXP
            }
        } 

        if ($PSWrecords.UPN.Count -gt 0) {
            $formlist = FormBase -w 480 -h ((($PSWrecords.UPN.Count) * 30) + 180) -text 'ACCESS'
            $they = 20
            $choices = @()
            foreach ($entry in $PSWrecords) {
                if ($they -eq 20) {
                    $isfirst = $true
                } else {
                    $isfirst = $false
                }
                $astring = "$($entry.UPN) on $($entry.APP)`n$($entry.EXP)"
                if ($entry.EXP -eq '(secret EXPIRED)') {
                    $choices += RadioButton -form $formlist -x 20 -y $they -w 450 -h 50 -checked $isfirst -text $astring -enabled $false
                } else {
                    $choices += RadioButton -form $formlist -x 20 -y $they -w 450 -h 50 -checked $isfirst -text $astring
                }
                $they += 50 
            }
            RETRYButton -form $formlist -x 280 -y ($they + 30) -w 120 -text "Update secret" | Out-Null
            OKButton -form $formlist -x 30 -y ($they + 30) -w 120 -text "Access" | Out-Null
            $resultButton = $formlist.ShowDialog()
            foreach ($item in $choices) {
                if ($item.Checked) {
                    $item.Text -match "^(.+) on (.+)`n" | Out-Null
                    $UPN_App = "$($matches[1])<<>>$($matches[2])"
                }
            }

            if ($resultButton -eq 'RETRY') {
                $pswout = PowerShell.exe -file $WhereIsMyWallet `
                    -ExtKey $WhereIsMyKey  `
                    -ExtScript $ascript  `
                    -ExtUsr $UPN_App  `
                    -ExtAction 'updateGraph'

            } elseif ($resultButton -eq 'OK') {
                $pswout = PowerShell.exe -file $WhereIsMyWallet `
                    -ExtKey $WhereIsMyKey  `
                    -ExtScript $ascript  `
                    -ExtUsr $UPN_App  `
                    -ExtAction 'getGraph'
            }

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
                        
        } else {
            [System.Windows.MessageBox]::Show("No record found on database",'INFO','Ok','Warning') | Out-Null
        }
    }
} else {
    [System.Windows.MessageBox]::Show("No PSWallet or related database found",'ABORTING','Ok','Error') | Out-Null
}

Write-Host -ForegroundColor Blue "$($idi.UPN)"
Write-Host -ForegroundColor Blue "$($idi.CLIENT)"
Write-Host -ForegroundColor Blue "$($idi.TENANT)"
Write-Host -ForegroundColor Blue "$($idi.SECRET)"
