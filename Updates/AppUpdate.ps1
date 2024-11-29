<#
Name......: AppUpdate.ps1
Version...: 24.11.2
Author....: Dario CORRADA

This script looks for installed apps and try to update them
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

# get app list
$stdout_file = "$env:USERPROFILE\Downloads\$(get-date -f yyMMdd-HHmmss)_WingetUpgrade.log"
Start-Process -Wait -FilePath "winget.exe" -ArgumentList 'list --upgrade-available --source winget' -NoNewWindow -RedirectStandardOutput $stdout_file
$upgradable = $false
$AppList = @{}
foreach ($newline in (Get-Content $stdout_file -Encoding UTF8)) {
    $matches = @()
    
    # collecting items
    if (($upgradable -eq $true) -and !($newline | Select-String -Pattern ("^Nome   ", "^Name   ", "^\-+$", "^[0-9]+ "))) {
        $newline -match "^(.{$colName})(.{$colId})(.{$colVer})(.+)$" | Out-Null
        if (([string]::IsNullOrEmpty($matches[1])) -or ([string]::IsNullOrEmpty($matches[2]))) {
            if (!([string]::IsNullOrEmpty($newline))) {
                Write-Host -ForegroundColor Yellow "WARNING: something doesn't work with the following string:"">>$newline<<`n"
                Pause
            }
        } else {
            $AppList[$matches[2].Trim()] = @{
                NAME    = $matches[1].Trim()
                VERSION = $matches[3].Trim()
                AVAIL   = $matches[4].Trim()
            }
        }
    }
    
    # looking for the header of the table
    if ($newline | Select-String -Pattern ('Nome   ', 'Name   ')) {
        $upgradable = $true
        $newline -match "^([A-Za-z]+)( +)([A-Za-z]+)( +)([A-Za-z]+)( +)" | Out-Null
        $colName = $matches[1].Length + $matches[2].Length
        $colId = $matches[3].Length + $matches[4].Length
        $colVer = $matches[5].Length + $matches[6].Length
    }
}

# show dialog
if ($upgradable -eq $true) {
    $adialog = FormBase -w 550 -h ((($AppList.Count-1) * 30) + 140) -text "UPGRADABLE APPS"
    $they = 10
    $selmods = @{}
    foreach ($item in ($AppList.Keys | Sort-Object)) {
        $desc = "$($AppList[$item].NAME)  $($AppList[$item].VERSION) >>> $($AppList[$item].AVAIL)"
        $selmods[$item] = CheckBox -form $adialog -checked $true -x 20 -y $they -w 520 -text $desc
        $they += 30
    }
    OKButton -form $adialog -x 200 -y ($they + 10) -text "Ok" | Out-Null
    $result = $adialog.ShowDialog()
} else {
    [System.Windows.MessageBox]::Show("$newline","THAT'S ALL FOLKS!",'Ok','Info')
    exit
}

# run upgrade
foreach ($currentId in $selmods.Keys) {
    if ($selmods[$currentId].Checked) {
        $currentArg = "upgrade --id $currentId"
        Write-Host -ForegroundColor Yellow "`n*** Upgrading $($selmods[$currentId].Text) ***"
        Start-Process -Wait -FilePath "winget.exe" -ArgumentList $currentArg -NoNewWindow
    }
}
[System.Windows.MessageBox]::Show("All selected apps have been processed","THAT'S ALL FOLKS!",'Ok','Info') | Out-Null
