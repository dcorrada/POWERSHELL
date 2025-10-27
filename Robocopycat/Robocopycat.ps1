<#
Name......: Robocopycat.ps1
Version...: 25.10.1
Author....: Dario CORRADA

This script performs a data mirroring using robocopy command: each subfolder 
found will be forked into a new robocopy job
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
        $ThirdParty = 'Ok'
    } catch {
        [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error') > $null
        Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
        exit
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                                    INIT
******************************************************************************* #>
Write-Host -ForegroundColor Cyan -NoNewline "Looking for source tree and replicate onto destination"

# paths
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") > $null
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.RootFolder = "MyComputer"
$foldername.ShowNewFolderButton = $false
$foldername.Description = "SOURCE FOLDER"
$foldername.ShowDialog() > $null
$SOURCEpath = $foldername.SelectedPath -replace '\\', '/'
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") > $null
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.RootFolder = "MyComputer"
$foldername.ShowNewFolderButton = $false
$foldername.Description = "DESTINATION FOLDER"
$foldername.ShowDialog() > $null
$DESTpath = $foldername.SelectedPath -replace '\\', '/'

# RETRIEVING SOURCE TREE
# the number of subfolders found will define the amount of independent robocopy
# jobs that will be parallelized
$sourceTree = @{}
# the deepSeek value specify how much deeply looking for subfolders recursively
$deepSeek = 10
$ErrorActionPreference= 'SilentlyContinue'
$sourceTree[$SOURCEpath] = 0
for ($i = 1; $i -lt $deepSeek; $i++) {
    foreach ($parent in $sourceTree.Keys) {
        if ($sourceTree[$parent] -eq ($i - 1)) {
            foreach ($child in Get-ChildItem -Path $parent) {
                if ($child.Mode -eq 'd-----') { # new subfolder
                    $astring = $parent + '/' + $child.Name
                    $sourceTree[$astring] = $i
                    Write-Host -NoNewline '.'
                }
            }
        }
    }
}
$ErrorActionPreference= 'Inquire'

# GENERATING DESTINATION TREE
$destTree = @{}
foreach ($item in $sourceTree.Keys) {
    $adest = $item -replace "$SOURCEpath", "$DESTpath"
    $destTree[$adest] = $sourceTree[$item]
    Write-Host -NoNewline '.'
}

for ($i = 1; $i -lt $deepSeek; $i++) {
    foreach ($item in $destTree.Keys) {
        if ($destTree[$item] -eq $i) {
            New-Item -Path $item -ItemType Directory | Out-Null
            Write-Host -NoNewline '.'
        }
    }
}
Write-Host -ForegroundColor Green " Done`n"


<#
+++++++++++++++++++++
+++  NOTE PER ME  +++
+++++++++++++++++++++

* Gestire un file exclude list per i path da escludere, verificare se in robocopy esiste una opzione exclude

* Rimuovere, dalla stringa del job di robocopy, l'opzione di girare ricorsivamente nelle sottocartelle 

* Togliere da GitHub il source tree di testing [unstable]/Robocopycat/RoboCopyCat_TEST
#>
