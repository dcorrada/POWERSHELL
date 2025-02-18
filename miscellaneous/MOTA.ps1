<#
Name......: MOTA.ps1
Version...: 25.02.1
Author....: Dario CORRADA

Questo script e' un proof of concept per ottenere una reportistica mensile del 
bilancio di licenze Microsoft 365 assegnate su recuperate. Lo script leggera' 
dal file Excel prodotto con lo script AssignedLicensesSDK.ps1. 

Comincera' leggendo il worksheet "Assigned_Licenses" per ottenere una lista 
preliminare di licenze assegnate in uno specifico mese.

Successivamente andra' a parsare le voci presenti nel worksheet "Orphaned": 
* cerchera' le voci annotate come "user no longer exists on tenant", o come
  "license dismissed for this user", per valutare se sia stata 
  effettivamente recuperata una licenza;
* cerchera' le voci annotate come "assigned license(s) to this user": in 
  questo caso occorrera' valutare sia stata assegnata una licenza, nel 
  periodo temporale di riferimento, e se poi sia stata successivamente 
  recuperata a causa di uno switch (ie passaggio da Basic a Standard) 
  oppure se l'utente sia cessato nello stesso periodo

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
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module -Name "$workdir\Modules\Forms.psm1"
        Import-Module ImportExcel
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'ImportExcel')) {
            Install-Module ImportExcel -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [ImportExcel] module: click Ok restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } else {
            [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error') > $null
            Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
            exit
        }
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'

[System.Windows.MessageBox]::Show("Everyday, well it's the same`nThat bong that's on the table starts to call my name`n`n[The Offspring]",'MOTA!','Ok','info') > $null


<# *******************************************************************************
                               FETCHING RAW DATA
******************************************************************************* #>
# looking for Excel reference file
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Open File"
$OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
$OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
$OpenFileDialog.ShowDialog() | Out-Null
$xlsx_file = $OpenFileDialog.filename

$Worksheet_list = Get-ExcelSheetInfo -Path $xlsx_file
$RawData = @{}

if ($Worksheet_list.Name -contains 'Assigned_Licenses') {
    Write-Host -NoNewline "Importing [Assigned_Licenses] worksheet..."
    foreach ($history in (Import-Excel -Path $xlsx_file -WorksheetName 'Assigned_Licenses')) {
        Write-Host -NoNewline '.'
        $aUSER      = $history.USRNAME
        $aDATE      = $history.TIMESTAMP | Get-Date -format "yyyy/MM/dd"
        $aLICENSE   = $history.LICENSE
        $aNOTE      = 'null'
        if ($aLICENSE -cne 'NONE') {
            if (!($RawData.ContainsKey($aUSER))) {
                $RawData["$aUSER"] = @{}
            }
            if (!($RawData["$aUSER"].ContainsKey($aDATE))) {
                ($RawData["$aUSER"])["$aDATE"] = @()
            }
            # $RawData["$aUSER"])["$aDATE"] sarà un array contenente stringhe in cui combino
            # licenza + note (ad es. "licenza>>note")
        }
    }
    Write-Host -ForegroundColor Green ' DONE'
} else {
    [System.Windows.MessageBox]::Show("The Excel file does not contain necessary data`n`n[$xlsx_file]",'ABORTING','Ok','Error') > $null
    Exit
}

if ($Worksheet_list.Name -contains 'Orphaned') {
    Write-Host -NoNewline "Importing [Orphaned] worksheet..."
    foreach ($history in (Import-Excel -Path $xlsx_file -WorksheetName 'Orphaned')) {
        Write-Host -NoNewline '.'
        $aUser = $history.USRNAME
    }
    Write-Host -ForegroundColor Green ' DONE'
} else {
    [System.Windows.MessageBox]::Show("Worksheet 'Orphaned not found in the Excel file`n`n[$xlsx_file]",'WARNING','Ok','Warning') > $null
}