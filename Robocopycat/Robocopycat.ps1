<#
Name......: Robocopycat.ps1
Version...: 25.10.1
Author....: Dario CORRADA

This script performs a data mirroring using robocopy command: each subfolder 
found will be forked into a new robocopy job.

In the exclude.list file you can edit a list of patterns that you would like 
to exclude. The script will try to guess folders and/or files to be excluded 
from the backup jobs.

For a complete reference on how robocopy works take a look to
https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
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
                                    INPUTS
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

# exclude list
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select exclude list file"
$OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
$OpenFileDialog.filename = 'exclude.list'
$OpenFileDialog.filter = 'Plain text file | *.*'
$OpenFileDialog.ShowDialog() | Out-Null
$ExcludeList = @{}
foreach ($item in (Get-Content $OpenFileDialog.filename)) {
    $ExcludeList[$item] = 'file'
}

# RETRIEVING SOURCE TREE
# the number of subfolders found will define the amount of independent robocopy
# jobs that will be parallelized
$sourceTree = @{}
# the deepSeek value specify how much deeply looking for subfolders recursively
$deepSeek = 10
$ErrorActionPreference= 'SilentlyContinue'
$sourceTree[$SOURCEpath] = 0
$deepestLevel = 0
for ($i = 1; $i -lt $deepSeek; $i++) {
    foreach ($parent in $sourceTree.Keys) {
        if ($sourceTree[$parent] -eq ($i - 1)) {
            foreach ($child in Get-ChildItem -Path $parent) {
                if ($child.Mode -eq 'd-----') { # new subfolder
                    $astring = $parent + '/' + $child.Name
                    $proceed = $true
                    foreach ($blackListed in $ExcludeList.Keys) {
                        if ($astring -cmatch "$blacklisted") {
                            $proceed = $false
                            $ExcludeList[$blackListed] = 'folder'
                        }
                    }
                    if ($proceed) {
                        $sourceTree[$astring] = $i
                        if ($i -gt $deepestLevel) {
                            $deepestLevel = $i
                        }
                        Write-Host -NoNewline '.'
                    }
                }
            }
        }
    }
}
$ErrorActionPreference= 'Inquire'

# each record of $jobArray couple a source path with a destination path,
# plus a value of how deep is nested such couple
$jobArray = @()
foreach ($item in $sourceTree.Keys) {
    $adest = $item -replace "$SOURCEpath", "$DESTpath"
    $arecord = @{
        SOURCE  = $item
        VALUE   = $sourceTree[$item]
        DEST    = $adest
    }
    $jobArray += $arecord
    Write-Host -NoNewline '.'
}

# GENERATING DESTINATION TREE
for ($i = 1; $i -lt $deepSeek; $i++) {
    foreach ($item in $jobArray) {
        if ($item.VALUE -eq $i) {
            New-Item -Path $item.DEST -ItemType Directory | Out-Null
            Write-Host -NoNewline '.'
        }
    }
}
Write-Host -ForegroundColor Green " Done`n"

<# *******************************************************************************
                                  DRY RUN
******************************************************************************* #>

$source_path = 'C:/Users/korda/Desktop/POWERSHELL/Robocopycat/RoboCopyCat_TEST'
$dest_path = 'C:/Users/korda/Downloads/TEST'
$logfile = 'C:/Users/korda/Downloads/robocopy.log'

# standard params adopted for
$stdParms = '/XJ /R:3 /NP /NDL /NC /NJH /NJS'

# extra params for those job at deepest level
$tailParms = '/E /MIR'

# extra params for dry runs
$dryParms = '/L /BYTES'

# extra params for excluding files
$excParms = '/XF'
foreach ($item in $ExcludeList.Keys) {
    if ($ExcludeList[$item] -eq 'file') {
        $excParms += " $item"
    }
}

$StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /E /MIR {3} /XF' -f $source_path, $dest_path, $logfile, $stdParms
Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList

<#
+++++++++++++++++++++
+++  NOTE PER ME  +++
+++++++++++++++++++++

* Gestire un file exclude list per i path da escludere
    gestire $blacklisted come se fosse un hash in cui flaggare se si tratta di un path o di un file (e quindi da aggiungere alla lista /XF nel job)

* Rimuovere, dalla stringa del job di robocopy, l'opzione di girare ricorsivamente nelle sottocartelle 

* Togliere da GitHub il source tree di testing [unstable]/Robocopycat/RoboCopyCat_TEST

* Una volta finiti i test spostare Data_Backup nel branch [tempus] essendo una versione legacy
#>
