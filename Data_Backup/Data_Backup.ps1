<#
Name......: Data_Backup.ps1
Version...: 21.07.1
Author....: Dario CORRADA

This script performs a complete user backup to an external disk or to a shared folder
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Data_Backup\\Data_Backup\.ps1$" > $null
$repopath = $matches[1]

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$repopath\Modules\Forms.psm1"

# temporary directory
$tmppath = 'C:\DATABACKUP_LOGS'
if (Test-Path $tmppath) {
    Remove-Item "$tmppath" -Recurse -Force
    Start-Sleep 2
}
New-Item -ItemType directory -Path $tmppath > $null

# select destination path
$AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
$Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
$OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
$OpenFileDialog.AddExtension = $false
$OpenFileDialog.CheckFileExists = $false
$OpenFileDialog.DereferenceLinks = $true
$OpenFileDialog.Filter = "Folders|`n"
$OpenFileDialog.Multiselect = $false
$OpenFileDialog.Title = "Select destination path"
$OpenFileDialogType = $OpenFileDialog.GetType()
$FileDialogInterfaceType = $Assembly.GetType('System.Windows.Forms.FileDialogNative+IFileDialog')
$IFileDialog = $OpenFileDialogType.GetMethod('CreateVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null)
$OpenFileDialogType.GetMethod('OnBeforeVistaDialog',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$IFileDialog)
[uint32]$PickFoldersOption = $Assembly.GetType('System.Windows.Forms.FileDialogNative+FOS').GetField('FOS_PICKFOLDERS').GetValue($null)
$FolderOptions = $OpenFileDialogType.GetMethod('get_Options',@('NonPublic','Public','Static','Instance')).Invoke($OpenFileDialog,$null) -bor $PickFoldersOption
$FileDialogInterfaceType.GetMethod('SetOptions',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$FolderOptions)
$VistaDialogEvent = [System.Activator]::CreateInstance($AssemblyFullName,'System.Windows.Forms.FileDialog+VistaDialogEvents',$false,0,$null,$OpenFileDialog,$null,$null).Unwrap()
[uint32]$AdviceCookie = 0
$AdvisoryParameters = @($VistaDialogEvent,$AdviceCookie)
$AdviseResult = $FileDialogInterfaceType.GetMethod('Advise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdvisoryParameters)
$AdviceCookie = $AdvisoryParameters[1]
$Result = $FileDialogInterfaceType.GetMethod('Show',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,[System.IntPtr]::Zero)
$FileDialogInterfaceType.GetMethod('Unadvise',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$AdviceCookie)
if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
    $FileDialogInterfaceType.GetMethod('GetResult',@('NonPublic','Public','Static','Instance')).Invoke($IFileDialog,$null)
}
$copiasu = $OpenFileDialog.FileName

# warn login on remote path (for shared folders)
if ($copiasu -match "^\\\\") {
    [System.Windows.MessageBox]::Show("Access on $copiasu with your credentials, then click Ok to continue",'WARNING','Ok','Warning')
}

if (!($copiasu -match "\\$")) {
    $copiasu = $copiasu + '\'
}

Write-Host -NoNewline "Checking paths to backup..."
$backup_list = @{} # variable in which will added paths to backup
$usrlist = @()
$root_path = 'C:\'
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null

# select user profiles
$userlist = Get-ChildItem C:\Users
$hsize = 175 + (25 * $userlist.Count)
$form_panel = FormBase -w 280 -h $hsize -text "USER FOLDERS"
Label -form $form_panel -x 10 -y 20 -text 'Select users to backup:' | Out-Null
$vpos = 50
$boxes = @()
foreach ($elem in $userlist) {
    if ($elem.Name -eq $env:USERNAME) {
        $boxes += CheckBox -form $form_panel -checked $true -enabled $false -x 20 -y $vpos -text $elem.Name
        $vpos += 25
    } else {
        $boxes += CheckBox -form $form_panel -checked $false -x 20 -y $vpos -text $elem.Name
        $vpos += 30
    }
}
$vpos += 20
OKButton -form $form_panel -x 80 -y $vpos -text "Ok" | Out-Null
$result = $form_panel.ShowDialog()
foreach ($box in $boxes) {
    if ($box.Checked -eq $true) {
        $usrname = $box.Text
        $usrlist += "$usrname"
    }
}

# load exclude list file
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select excluded path list file"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Plain text file | *.*'
$OpenFileDialog.ShowDialog() | Out-Null
$excludefile = $OpenFileDialog.filename

# load allow list file
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select allowed path list file"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Plain text file | *.*'
$OpenFileDialog.ShowDialog() | Out-Null
$allowfile = $OpenFileDialog.filename

# adding specific allowed paths
[string[]]$allow_list = Get-Content -Path $allowfile
foreach ($item in $allow_list) {
    if ($item -match '\$username') {
        foreach ($usr in $usrlist) {
            $replaced = $item -replace ('\$username', $usr)            
            $full_path = $root_path + $replaced
            if (Test-Path $full_path) {
                $CommonRobocopyParams = '/E /XJ /R:0 /W:1 /MT:64 /NP /NDL /NC /BYTES /NJH /NJS'
                $StagingLogPath = $tmppath + '\test.log'
                $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $full_path, $StagingLogPath, $CommonRobocopyParams
                Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
                $StagingContent = Get-Content -Path $StagingLogPath
                $TotalFileCount = $StagingContent.Count
                $backup_list[$replaced] = $TotalFileCount
                Remove-Item $StagingLogPath -Force
                Write-Host -NoNewline "."
            }
        }
    } else {
        $full_path = $root_path + $item
        if (Test-Path $full_path) {
                $CommonRobocopyParams = '/E /XJ /R:0 /W:1 /MT:64 /NP /NDL /NC /BYTES /NJH /NJS'
                $StagingLogPath = $tmppath + '\test.log'
                $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $full_path, $StagingLogPath, $CommonRobocopyParams
                Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
                $StagingContent = Get-Content -Path $StagingLogPath
                $TotalFileCount = $StagingContent.Count
                $backup_list[$item] = $TotalFileCount
                Remove-Item $StagingLogPath -Force
                Write-Host -NoNewline "."
        }        
    }
}

# adding default paths
[string[]]$exclude_list = Get-Content -Path $excludefile
$exclude_list += 'Users' # users' folders will be evaluated one by one
# adding root paths
$rooted = Get-ChildItem $root_path -Attributes D
foreach ($candidate in $rooted) {
    $decision = $true
    foreach ($item in $exclude_list) {
        if ($item -eq $candidate) {
            $decision = $false
        }
    }
    if ($decision) {
        $full_path = $root_path + $candidate
        $CommonRobocopyParams = '/E /XJ /R:0 /W:1 /MT:64 /NP /NDL /NC /BYTES /NJH /NJS'
        $StagingLogPath = $tmppath + '\test.log'
        $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $full_path, $StagingLogPath, $CommonRobocopyParams
        Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
        $StagingContent = Get-Content -Path $StagingLogPath
        $TotalFileCount = $StagingContent.Count
        $backup_list[$candidate] = $TotalFileCount
        Remove-Item $StagingLogPath -Force
        Write-Host -NoNewline "."
    }
}

# adding users' folders
foreach ($usr in $usrlist) {
    $root_usr = $root_path + 'Users\' + $usr + '\'
    $rooted = Get-ChildItem $root_usr -Attributes D
    $exclude_usrlist = $exclude_list -replace ('\$username', $usr)
    foreach ($candidate in $rooted) {
        $candidate = 'Users\' + $usr + '\' + $candidate
        $decision = $true
        foreach ($item in $exclude_usrlist) {
            if ($item -eq $candidate) {
                $decision = $false
            } elseif ($candidate -eq 'AppData') { # such folder will be parsed specifically
                $decision = $false
            }
        }
        if ($decision) {
            $full_path = $root_path + $candidate
            $CommonRobocopyParams = '/E /XJ /R:0 /W:1 /MT:64 /NP /NDL /NC /BYTES /NJH /NJS'
            $StagingLogPath = $tmppath + '\test.log'
            $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $full_path, $StagingLogPath, $CommonRobocopyParams
            Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
            $StagingContent = Get-Content -Path $StagingLogPath
            $TotalFileCount = $StagingContent.Count
            $backup_list[$candidate] = $TotalFileCount
            Remove-Item $StagingLogPath -Force
            Write-Host -NoNewline "."
        }
    }  
}

# adding users' AppData hidden folders
foreach ($usr in $usrlist) {
    foreach ($subfolder in ('Local', 'LocalLow', 'Roaming')) {
        $root_usr = $root_path + 'Users\' + $usr + '\AppData\' + $subfolder + '\'
        $rooted = Get-ChildItem $root_usr -Attributes D
        $exclude_usrlist = $exclude_list -replace ('\$username', $usr)
        foreach ($candidate in $rooted) {
            $candidate = 'Users\' + $usr + '\AppData\' + $subfolder + '\' + $candidate
            $decision = $true
            foreach ($item in $exclude_usrlist) {
                if ($item -eq $candidate) {
                    $decision = $false
                }
            }
            if ($decision) {
                $full_path = $root_path + $candidate
                $CommonRobocopyParams = '/E /XJ /R:0 /W:1 /MT:64 /NP /NDL /NC /BYTES /NJH /NJS'
                $StagingLogPath = $tmppath + '\test.log'
                $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $full_path, $StagingLogPath, $CommonRobocopyParams
                Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
                $StagingContent = Get-Content -Path $StagingLogPath
                $TotalFileCount = $StagingContent.Count
                $backup_list[$candidate] = $TotalFileCount
                Remove-Item $StagingLogPath -Force
                Write-Host -NoNewline "."
            }
        }
    }  
}

Write-Host -ForegroundColor Green " DONE"

Write-Host -ForegroundColor Yellow "`nThe following paths will be backupped onto [$copiasu]`n"
$TotalFileToBackup = 0
foreach ($item in ($backup_list.Keys | Sort-Object)) {
    $string = $root_path + $item
    $files = $backup_list[$item]
    $TotalFileToBackup += $files
    Write-Host -NoNewline "$string "
    Write-Host -ForegroundColor Cyan "$files file(s) to backup"
}
Write-Host -ForegroundColor Yellow "`nTOTAL $TotalFileToBackup file(s) to backup"

$answ = [System.Windows.MessageBox]::Show("Do you want to proceed?",'PROCEED','YesNo','Info')
if ($answ -eq "No") {    
    Exit
}

# create destination folder
try {
    $copiasu = $copiasu + $env:USERNAME + '_on_' + $env:COMPUTERNAME
    if (!(Test-Path $copiasu)) {
        New-Item -ItemType directory -Path $copiasu > $null
    }
}
catch {
    [System.Windows.MessageBox]::Show("Unable to create $copiasu",'ERROR','Ok','Error') > $null
    Exit
}

# backup job block
Write-Host " "
$RoboCopyBlock = {
    param($final_path, $prefix, $logpath)
    $filename = $final_path -replace ('\\','-')
    if (Test-Path "$logpath\ROBOCOPY_$filename.log" -PathType Leaf) {
        Remove-Item  "$logpath\ROBOCOPY_$filename.log" -Force
    }
    New-Item -ItemType file "$logpath\ROBOCOPY_$filename.log" > $null
    $source = 'C:\' + $final_path
    $dest = $prefix + '\' + $final_path

    # for the options see https://superuser.com/questions/814102/robocopy-command-to-do-an-incremental-backup
    $opts = ("/E", "/XJ", "/R:5", "/W:10", "/NP", "/NDL", "/NC", "/NJH", "/ZB", "/MIR", "/LOG+:$logpath\ROBOCOPY_$filename.log")
    if ($prefix -match "^\\\\") {
        $opts += '/COMPRESS'
    }  
    $cmd_args = ($source, $dest, $opts)
    robocopy @cmd_args
}

# launch multithreaded backup jobs
foreach ($folder in $backup_list.Keys) {
    Start-Job $RoboCopyBlock -Name $folder -ArgumentList $folder, $copiasu, $tmppath > $null
}

# progress bar
$form_bar = New-Object System.Windows.Forms.Form
$form_bar.Text = "TRANSFER RATE"
$form_bar.Size = New-Object System.Drawing.Size(600,200)
$form_bar.StartPosition = 'CenterScreen'
$form_bar.Topmost = $true
$form_bar.MinimizeBox = $false
$form_bar.MaximizeBox = $false
$form_bar.FormBorderStyle = 'FixedSingle'
$font = New-Object System.Drawing.Font("Arial", 12)
$form_bar.Font = $font
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(20,20)
$label.Size = New-Object System.Drawing.Size(550,30)
$form_bar.Controls.Add($label)
$bar = New-Object System.Windows.Forms.ProgressBar
$bar.Style="Continuous"
$bar.Location = New-Object System.Drawing.Point(20,70)
$bar.Maximum = 101
$bar.Size = New-Object System.Drawing.Size(550,30)
$form_bar.Controls.Add($bar)
$form_bar.Show() | out-null

# Waiting for jobs completed
While (Get-Job -State "Running") {
    Clear-Host
    Write-Host -ForegroundColor Blue "*** BACKUP ***"
    
    $directoryInfo = Get-ChildItem $tmppath | Measure-Object
    if ($directoryInfo.count -eq 0) {
        Write-Host -ForegroundColor Cyan "Waiting jobs to start..."
    } else {
        $ActiveJobs = 0
        $TotalFilesCopied = 0
        foreach ($item in ($backup_list.Keys | Sort-Object)) {
            $full_path = "$root_path" + "$item"
            $string = $item -replace ('\\','-')
            $logfile = "$tmppath\ROBOCOPY_$string.log"
            $FilesCopied = 0
            $amount = $backup_list[$item]
            if (Test-Path "$logfile" -PathType Leaf) {
                $chomp = Get-Content -Path $logfile
                foreach ($newline in $chomp) {
                    $newline -match "(^\s+\d+)" > $null
                    if ($Matches[1]) {
                        $FilesCopied ++
                        $Matches[1] = $null
                    }
                }
                $ErrorActionPreference= 'SilentlyContinue'
                $output = [system.String]::Join(" ", $chomp)
                $ErrorActionPreference= 'Inquire'
                if ($output -match "-------------------------------------------------------------------------------") {
                    $donothing = 1
                    # Write-Host -NoNewline "[$full_path] "
                    # Write-Host -ForegroundColor Green "backupped!"
                } else {
                    $ActiveJobs ++
                    $FilesCopied = $FilesCopied - 1
                    if ($FilesCopied -lt 0) {
                        $FilesCopied = 0
                    }
                    Write-Host -NoNewline "[$full_path] "
                    Write-Host -ForegroundColor Cyan "$FilesCopied out of $amount file(s) copied"
                    $LastLine = Get-Content -Path $logfile -Tail 1
                    
                    $Matches[1] = $null
                    $LastLine -match "(^\s+\d+)" > $null
                    if ($Matches[1]) {
                        Write-Host -ForegroundColor Yellow ">>> $LastLine `n"
                    } else {
                        Write-Host -ForegroundColor Red "some error occurs, see $logfile `n"
                    }
                }
            }
            $TotalFilesCopied += $FilesCopied
        }
        if ($ActiveJobs -lt 1) {
            Write-Host -ForegroundColor Cyan "Checking backup..."                
        }      
        
        $percent = ($TotalFilesCopied / $TotalFileToBackup)*100
        if ($percent -gt 100) {
            $percent = 100
        }
        $formattato = '{0:0.0}' -f $percent
        [int32]$progress = $percent
        Write-Host -ForegroundColor Yellow  "`nTOTAL PROGRESS: $formattato%"

        $label.Text = "Progress: $formattato% - $TotalFilesCopied out of $TotalFileToBackup copied"
        if ($progress -ge 100) {
            $bar.Value = 100
        } else {
            $bar.Value = $progress
        }

        # refreshing the progress bar
        [System.Windows.Forms.Application]::DoEvents()    
    }
    Start-Sleep -Milliseconds 500
}

$form_bar.Close()

$joblog = Get-Job | Receive-Job # get job output
Remove-Job * # Cleanup

# Size check
Write-Host " "
Write-Host -NoNewline "Size check..."
foreach ($folder in $backup_list.Keys) {
    $source = $root_path + $folder
    $dest = $copiasu + '\' + $folder
    $CommonRobocopyParams = '/E /XJ /R:0 /W:1 /MT:64 /NP /NDL /NC /BYTES /NJH'
    $StagingLogPath = $tmppath + '\test.log'

    $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $source, $StagingLogPath, $CommonRobocopyParams
    Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
    $StagingContent = Get-Content -Path $StagingLogPath
    $ErrorActionPreference= 'SilentlyContinue'
    $output = [system.String]::Join(" ", $StagingContent)
    $ErrorActionPreference= 'Inquire'
    $output -match "Byte:\s+(\d+)\s+\d+" > $null
    $source_size = $Matches[1]
    Remove-Item $StagingLogPath -Force

    $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $dest, $StagingLogPath, $CommonRobocopyParams
    Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
    $StagingContent = Get-Content -Path $StagingLogPath
    $ErrorActionPreference= 'SilentlyContinue'
    $output = [system.String]::Join(" ", $StagingContent)
    $ErrorActionPreference= 'Inquire'
    $output -match "Byte:\s+(\d+)\s+\d+" > $null
    $dest_size = $Matches[1]
    Remove-Item $StagingLogPath -Force

    Write-Host -NoNewline "[$folder] $source_size/$dest_size "

    if ($dest_size -lt $source_size) { # backup job failed
        Write-Host -ForegroundColor Red "DIFF"

        $whatif = [System.Windows.MessageBox]::Show("Copy of $folder failed.`nRelaunch backup job?",'ERROR','YesNo','Error')
        if ($whatif -eq "Yes") {
            $opts = ("/E", "/ZB", "/NP", "/W:5")
            $cmd_args = ($source, $dest, $opts)    
            Write-Host -ForegroundColor Yellow "RETRY: copy of $folder in progress..."
            Start-Sleep 3
            robocopy @cmd_args
            $whatif = [System.Windows.MessageBox]::Show("Backup of $folder is ok?",'CONFIRM','YesNo','Info')                
            if ($whatif -eq "No") {
                [System.Windows.MessageBox]::Show("Backup $folder manually",'CONFIRM','Ok','Info') > $null
            }
        }
    } else {
        Write-Host -ForegroundColor Green "OK"
    }    
}
$ErrorActionPreference= 'Inquire'
Write-Host -ForegroundColor Green " DONE"

# check folder attributes
Write-Host " "
Write-Host -NoNewline "Check attributes..."
foreach ($folder in $backup_list.Keys) {
    $dest = $copiasu + '\' + $folder
    attrib -s -h $dest
}
Write-Host -ForegroundColor Green " DONE"

# file backup from Users folder
foreach ($usr in $usrlist) {
    $prefix = $root_path + 'Users\' + $usr
    Write-Host -NoNewline "Copying files in $prefix..."
    $userfiles = Get-ChildItem "$prefix" -Attributes A
    foreach ($afile in $userfiles) {
        Copy-Item "$prefix\$afile" -Destination "$copiasu\Users\$usr" -Force > $null
    }
    Write-Host -ForegroundColor Green " DONE"

    Write-Host -NoNewline "Copying files in $prefix\AppData\Local..."
    $userfiles = Get-ChildItem "$prefix\AppData\Local" -Attributes A
    foreach ($afile in $userfiles) {
        Copy-Item "$prefix\AppData\Local\$afile" -Destination "$copiasu\Users\$usr\AppData\Local" -Force > $null
    }
    Write-Host -ForegroundColor Green " DONE"
}

# cleaning temporary
$answ = [System.Windows.MessageBox]::Show("Backup finished. Delete log files?",'END','YesNo','Info')
if ($answ -eq "Yes") {
    Remove-Item "$tmppath" -Recurse -Force
}
