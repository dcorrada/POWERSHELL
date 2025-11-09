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
                                DEFINING INPUTS
******************************************************************************* #>
# define source path
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



<# *******************************************************************************
                          BUILDING ARRAY OF JOB
******************************************************************************* #>
# create log folder
$logPath = 'C:\ROBOCOPYCAT_TEMP'
if (Test-Path $logPath) {
    Remove-Item $logPath -Recurse -Force
}
New-Item $logPath -ItemType Directory | Out-Null


Write-Host -ForegroundColor Cyan -NoNewline "`nExploring source tree"
$deepestLevel = 0
$jobArray = @{}
$jobCounter = 0
foreach ($item in $children) {
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
            $ajobname = 'ROBOCOP-' + ('{0:d4}' -f $jobCounter)
            $jobArray[$ajobname] = @{
                SOURCE_PATH = $item.FullName
                DEST_PATH   = $null
                PARAMS      = "/XJ /R:3 /W:10 /ZB /NP /NDL /NC /BYTES /LOG+:$logPath\$ajobname.log"
                LEVEL       = $SubLevels
                STATUS      = 'queued'
                
            }
            Write-Host -NoNewline '.' 

            # update the deepest level reached amomng all jobs
            if ($SubLevels -gt $deepestLevel) {
                $deepestLevel = $SubLevels
            }
        }
    }
}
Write-Host -ForegroundColor Green ' Done'



<# *******************************************************************************
                              DEFINING OUTPUTS
******************************************************************************* #>
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") > $null
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.RootFolder = "MyComputer"
$foldername.ShowNewFolderButton = $false
$foldername.Description = "DESTINATION FOLDER"
$foldername.ShowDialog() > $null
$DESTpath = $foldername.SelectedPath -replace '\\', '/'

Write-Host -ForegroundColor Cyan -NoNewline "`nGenerating destination tree"
for ($i = 1; $i -le $deepSeek; $i++) {
    foreach ($item in $jobArray.Keys) {
        if ($jobArray[$item].LEVEL -eq $i) {
            Write-Host -NoNewline '.'
            $apath = $jobArray[$item].SOURCE_PATH -replace '\\', '/'
            $apath -match "^$SOURCEpath(.*)$" | Out-Null
            $jobArray[$item].DEST_PATH = ("$DESTpath" + "$($matches[1])") -replace '/', '\'
            New-Item -Path $jobArray[$item].DEST_PATH -ItemType Directory | Out-Null
        }
    }
}
Write-Host -ForegroundColor Green ' Done'



<# *******************************************************************************
                              TUNING PARAMS
******************************************************************************* #>
# for those paths that define the lowest level of the destination tree recursive 
# setting options will be added
Write-Host -ForegroundColor Cyan -NoNewline "`nSelecting recursive jobs"
foreach ($item in $jobArray.Keys) {
    if ($jobArray[$item].LEVEL -eq $deepestLevel) {
        $jobArray[$item].PARAMS += ' /E /MIR'
        Write-Host -NoNewline '.'
    }
}
Write-Host -ForegroundColor Green ' Done'

# adding exclusionns
$excParms = ''
# extra params for excluding files
if ($ExcludeList.FILES.Count -gt 0) {
    $excParms += ' /XF'
    foreach ($item in $ExcludeList.FILES) {
        $string = ' "' + $item + '"'
        $excParms += $string
    }    
}
# extra params for excluding fdolders
if ($ExcludeList.FOLDERS.Count -gt 0) {
    $excParms += ' /XD'
    foreach ($item in $ExcludeList.FOLDERS) {
        $string = ' "' + $item + '"'
        $excParms += $string
    }    
}
foreach ($item in $jobArray.Keys) {
    $jobArray[$item].PARAMS += $excParms
}

# dry run
$answ = [System.Windows.MessageBox]::Show("Would you perform a simulation `ninstead of effective data transfer?",'DRY RUN','YesNo','Info')
if ($answ -eq "Yes") { 
    Write-Host -ForegroundColor Cyan -NoNewline "`nCleaning dest paths"
    for ($i = $deepSeek; $i -ge 1; $i--) {
        foreach ($item in $jobArray.Keys) {
            if ($jobArray[$item].LEVEL -eq $i) {
                $jobArray[$item].PARAMS += ' /L'
                Remove-Item -Path $jobArray[$item].DEST_PATH -Recurse -Force | Out-Null
                Write-Host -NoNewline '.'
            }
        }    
    }
    Write-Host -ForegroundColor Green ' Done'
}



<# *******************************************************************************
                                 JOB RUN
******************************************************************************* #>

<# NOTE PER ME DA DEBUGGARE
Sembra che parta il primo blocco di runs e poi finsice tutto
In realtà nemmeno il primo blocco di runs parte (altrimenti avrei dei log files)
#>


$RoboCopyBlock = {
    params($source_path, $dest_path, $opts)
    $argstring = "$source_path $dest_path $opts" 
    Start-Process -Wait -FilePath Robocopy.exe -ArgumentList $argstring
}

# this variable defineshow many job will run simultaneuosly
$ConcurrentRuns = 8

$RunningJobs = 0
do {

    foreach ($item in $jobArray.Keys) {
        if (($jobArray[$item].STATUS -eq 'queued') -and ($RunningJobs -lt $ConcurrentRuns)) {
            $jobArray[$item].STATUS = 'started'
            Start-Job $RoboCopyBlock -Name $item -ArgumentList $jobArray[$item].SOURCE_PATH, $jobArray[$item].DEST_PATH, $jobArray[$item].PARAMS | Out-Null
            $RunningJobs++
        }
    }

    $RunningJobs = 0
    $jobSnapshot = Get-Job
    foreach ($photo in $jobSnapshot) {
        if ($photo.State -cne 'Completed') {
            $RunningJobs++
        }
    }
    
    $StillAlive = 0
    Clear-Host
    Write-Host -ForegroundColor Yellow "*** RUNNING JOBS ***"
    foreach ($item in $jobArray.Keys) {
        $ErrorActionPreference= 'Stop'
        try {
            if ((Get-Job -Name $item).State -eq 'Running') {
                Write-Host -ForegroundColor Green "[$item] From <$($jobArray[$item].SOURCE_PATH)> to <$DESTpath>"
            }
        }
        catch {
            <# do nothing #>
        }
        $ErrorActionPreference= 'Inquire'
    }

    Write-Host -ForegroundColor Blue -NoNewline "`n`nPENDING JOBS "
    $pending_counter = 0
    foreach ($item in $jobArray.Keys) {
        $ErrorActionPreference= 'Stop'
        try {
            if ((Get-Job -Name $item).State -ne 'Completed') {
                $StillAlive++
            }
        }
        catch {
            $pending_counter++
        }
        $ErrorActionPreference= 'Inquire'
    }
    Write-Host "$pending_counter"

    Start-Sleep -Milliseconds 2000
} while ($StillAlive -gt 0)


<# *******************************************************************************
                                 JOB RUN
******************************************************************************* #>
$answ = [System.Windows.MessageBox]::Show("Would you keep log files stored in [$logpath]?",'DRY RUN','YesNo','Info')
if ($answ -eq "No") { 
    Remove-Item -Path $logPath -Recurse -Force | Out-Null
}