<#
Name......: Robocopycat.ps1
Version...: 25.11.1
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
Write-Host -ForegroundColor Cyan -NoNewline "`nDefining for source tree"

# paths
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") > $null
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.RootFolder = "MyComputer"
$foldername.ShowNewFolderButton = $false
$foldername.Description = "SOURCE FOLDER"
$foldername.ShowDialog() > $null
$SOURCEpath = $foldername.SelectedPath -replace '\\', '/'

# initial params for define the amount of jobs
$DeepForm = FormBase -w 450 -h 300 -text 'SAERCHING LEVELS'
Label -form $DeepForm -x 80 -y 10 -w 150 -text 'RECURSE LEVEL' | Out-Null
$recurselevel = Slider -form $DeepForm -x 20 -y 35 -min 1 -max 8 -defval 2
$recurselevelLabel = Label -form $DeepForm -x 225 -y 40 -w 20 -text $recurselevel.Value
$recurselevel.add_ValueChanged({
    $SliderValue = $recurselevel.Value
    $TextString =  $SliderValue
    $recurselevelLabel.Text = $TextString
})
$addhidden = CheckBox -form $DeepForm -x 260 -y 15 -text 'Include hidden folders' 
$addexclude = CheckBox -form $DeepForm -x 260 -y 45 -text 'Import exclude list' 
Label -form $DeepForm -x 30 -y 90 -w 350 -h 100 -text @"
Recurse level value will define how many levels of nested subfolders to consider as separate backup jobs.

Low values are suggested if you have to backup several folder at the same tree level, otherwise consider higher value for few folders heavy nested.
"@ | Out-Null
OKButton -form $DeepForm -x 150 -y 190 -w 120 -text "Proceed" | Out-Null
$resultButton = $DeepForm.ShowDialog()
if ($addhidden.Checked) {
    $children = Get-ChildItem -Path $SOURCEpath -Directory -Recurse -Force
} else {
    $children = Get-ChildItem -Path $SOURCEpath -Directory -Recurse
}
$DeepSeek = $recurselevel.Value

# exclude list
$ExcludeList = @{
    FILES   = @()
    FOLDERS = @()
}
if ($addexclude.Checked) {
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Title = "Select exclude list file"
    $OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
    $OpenFileDialog.filename = 'exclude.list'
    $OpenFileDialog.filter = 'Plain text file | *.*'
    $OpenFileDialog.ShowDialog() | Out-Null
    foreach ($item in (Get-Content $OpenFileDialog.filename)) {
        Write-Host -NoNewline '.'
        if ($item -match "^\+\s") {
            $item -match "^\+\s(.+)$" | Out-Null
            $ExcludeList.FOLDERS += $matches[1]
        } elseif ($item -match "^\-\s") {
            $item -match "^\-\s(.+)$" | Out-Null
            $ExcludeList.FILES += $matches[1]
        }
    }
}

$jobArray = @()
$deepestLevel = 0
$jobCounter = 0
foreach ($item in $children) {
    Write-Host -NoNewline '.' 
    $NewPath = $item.FullName -replace '\\', '/'
    $SubLevels = 0
    while ($NewPath -cne $SOURCEpath) {
        $NewPath = (Split-Path $NewPath -Parent) -replace '\\', '/'
        $SubLevels ++
    }
    if ($SubLevels -le $DeepSeek) {
        $includeRecord = $true
        if ($ExcludeList.FOLDERS.Count -gt 0) {
            foreach ($pattern in $ExcludeList.FOLDERS) {
                if (($item.FullName -replace '\\', '/') -match ($pattern -replace '\\', '/')) {
                    $includeRecord = $false
                }
            }
        }
        
        if ($includeRecord) {
            $jobCounter++
            $arecord = @{
                JOBNANE = 'ROBOCOP-' + ('{0:d3}' -f $jobCounter)
                SOURCE  = $item.FullName
                LEVEL   = $SubLevels
                DEST    = $null
                TOTFILE = 0
            }
            $jobArray += $arecord

            # update the deepest level reached amomng all jobs
            if ($SubLevels -gt $deepestLevel) {
                $deepestLevel = $SubLevels
            }
        }
    }
}
Write-Host -ForegroundColor Green ' Done'

# generating destination tree
Write-Host -ForegroundColor Cyan -NoNewline "`nSetting destination path"
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") > $null
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.RootFolder = "MyComputer"
$foldername.ShowNewFolderButton = $false
$foldername.Description = "DESTINATION FOLDER"
$foldername.ShowDialog() > $null
$DESTpath = $foldername.SelectedPath -replace '\\', '/'
for ($i = 1; $i -le $deepSeek; $i++) {
    foreach ($item in $jobArray) {
        if ($item.LEVEL -eq $i) {
            Write-Host -NoNewline '.'
            $apath = $item.SOURCE -replace '\\', '/'
            $apath -match "^$SOURCEpath(.*)$" | Out-Null
            $item.DEST = ("$DESTpath" + "$($matches[1])") -replace '/', '\'
            New-Item -Path $item.DEST -ItemType Directory | Out-Null
        }
    }
}
Write-Host -ForegroundColor Green ' Done'


<# *******************************************************************************
                                  DRY RUN
******************************************************************************* #>
Write-Host -ForegroundColor Cyan -NoNewline "`nPerforming a simulation of data transfer"

# create log folder
if (Test-Path 'C:\ROBOCOPYCAT_TEMP') {
    Remove-Item 'C:\ROBOCOPYCAT_TEMP' -Force
}
New-Item 'C:\ROBOCOPYCAT_TEMP' -ItemType Directory | Out-Null

# standard parameters (please note /L is specific for dry runs)
$stagingParms = '/XJ /R:3 /NP /NDL /NC /NJH /NJS /BYTES /LOG:"C:\ROBOCOPYCAT_TEMP\test.log" /L'
# extra params for excluding files
if ($ExcludeList.FILES.Count -gt 0) {
    $excParms = ' /XF'
    foreach ($item in $ExcludeList.FILES) {
        $string = ' "' + $item + '"'
        $excParms += $string
    }    
    $stagingParms += $excParms
}
# extra params for excluding fdolders
if ($ExcludeList.FOLDERS.Count -gt 0) {
    $excParms = ' /XD'
    foreach ($item in $ExcludeList.FOLDERS) {
        $string = ' "' + $item + '"'
        $excParms += $string
    }    
    $stagingParms += $excParms
}


foreach ($dryjob in $jobArray) {
    Write-Host -NoNewline '.'
    if ($dryjob.LEVEL -eq $deepestLevel) {
        # extra params for those job at deepest level
        $StagingArgumentList = '"{0}" c:\fakepath {1} /E /MIR' -f $dryjob.SOURCE, $stagingParms
    } else {
        $StagingArgumentList = '"{0}" c:\fakepath {1}' -f $dryjob.SOURCE, $stagingParms
    }
    Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
    $StagingContent = Get-Content -Path "C:\ROBOCOPYCAT_TEMP\test.log"
    $dryjob.TOTFILE = $StagingContent.Count
    Remove-Item "C:\ROBOCOPYCAT_TEMP\test.log" -Force
}
Write-Host -ForegroundColor Green ' Done'

# output a summary
Clear-Host
Write-Host -ForegroundColor Yellow @"
*** READY TO GO SUMMARY ***

FILES   SOURCE PATH
-------------------------------------------------------------------------------
"@
foreach ($item in $jobArray) {
    $anumber = '{0:d5}' -f $item.TOTFILE
    Write-Host -ForegroundColor Blue -NoNewline "$anumber"
    if ($item.LEVEL -eq $deepestLevel) {
        Write-Host -ForegroundColor Red -NoNewline '+'
        Write-Host -ForegroundColor Cyan "  $($item.SOURCE)"
    } else {
        Write-Host -ForegroundColor Cyan "   $($item.SOURCE)"
    }
}
Write-Host -ForegroundColor Yellow "-------------------------------------------------------------------------------"

<# *******************************************************************************
                                  JOB RUN
******************************************************************************* #>
$answ = [System.Windows.MessageBox]::Show("Do you want to proceed to data transfer?",'STARTJOB','YesNo','Info')
if ($answ -eq "Yes") {
<# *** TO DO *** 

NOTE A MARGINE
* Rimuovere, dalla stringa del job di robocopy, l'opzione di girare 
  ricorsivamente nelle sottocartelle (a meno delle cartelle a piu basso livello)

* Una volta finiti i test spostare Data_Backup nel branch [tempus] essendo una
  versione legacy/deprecated
#>
} else {
    Remove-Item 'C:\ROBOCOPYCAT_TEMP' -Force
    Write-Host -ForegroundColor Cyan -NoNewline "`nRemoving destination paths"
    foreach ($item in $jobArray) {
        Write-Host -NoNewline '.'
        if ($item.LEVEL -eq 1) {
            Remove-Item $item.DEST -Recurse -Force
        }
    }
    Write-Host -ForegroundColor Green " Done"
    Start-Sleep -Milliseconds 1500
}

<#
+++++++++++++++++++++
+++  NOTE PER ME  +++
+++++++++++++++++++++

* Gestire un file exclude list per i path da escludere
    gestire $blacklisted come se fosse un hash in cui flaggare se si tratta di un path o di un file (e quindi da aggiungere alla lista /XF nel job)


#>
