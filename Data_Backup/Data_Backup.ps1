<#
Name......: Data_Backup.ps1
Version...: 20.12.1
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

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$repopath\Modules\Forms.psm1"

# temporary directory
$tmppath = 'C:\DATABACKUP_LOGS'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
}

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

# create destination folder
try {
    $copiasu = $copiasu + 'BACKUP_of_' + $env:USERNAME + '_on_' + (Get-Date -Format "yyyy-MM-dd_HH.mm")
    New-Item -ItemType directory -Path $copiasu > $null
}
catch {
    [System.Windows.MessageBox]::Show("Unable to create $copiasu",'ERROR','Ok','Error') > $null
    Exit
}

Write-Host -NoNewline "Searching paths to backup..."
$backup_list = @{} # variable in which will added paths to backup
$usrlist = @()
$root_path = 'C:\'
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null

# select user profiles
$userlist = Get-ChildItem C:\Users
$hsize = 200 + (30 * $userlist.Count)
$form_panel = FormBase -w 300 -h $hsize -text "USER FOLDERS"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(200,30)
$label.Text = "Select users to backup:"
$form_panel.Controls.Add($label)
$vpos = 50
$boxes = @()
foreach ($elem in $userlist) {
    if ($elem.Name -eq $env:USERNAME) {
        $boxes += CheckBox -form $form_panel -checked $true -enabled $false -x 20 -y $vpos -text $elem.Name
        $vpos += 30
    } else {
        $boxes += CheckBox -form $form_panel -checked $false -x 20 -y $vpos -text $elem.Name
        $vpos += 30
    }
}
$vpos += 20
OKButton -form $form_panel -x 90 -y $vpos -text "Ok"
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
                $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
                $output = [system.String]::Join(" ", $output)
                $output -match "Byte:\s+(\d+)\s+\d+" > $null
                $size = $Matches[1]
                $backup_list[$replaced] = $size
                Write-Host -NoNewline "."
            }
        }
    } else {
        $full_path = $root_path + $item
        if (Test-Path $full_path) {
            $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
            $output = [system.String]::Join(" ", $output)
            $output -match "Byte:\s+(\d+)\s+\d+" > $null
            $size = $Matches[1]
            $backup_list[$item] = $size
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
        $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Byte:\s+(\d+)\s+\d+" > $null
        $size = $Matches[1]
        $backup_list[$candidate] = $size
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
            }
        }
        if ($decision) {
            $full_path = $root_path + $candidate
            $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
            $output = [system.String]::Join(" ", $output)
            $output -match "Byte:\s+(\d+)\s+\d+" > $null
            $size = $Matches[1]
            $backup_list[$candidate] = $size
            Write-Host -NoNewline "."
        }
    }  
}

Write-Host -ForegroundColor Green " DONE"

Write-Host -ForegroundColor Yellow "`nThe following paths will be backupped onto [ $copiasu ]"
foreach ($item in ($backup_list.Keys | Sort-Object)) {
    $string = $root_path + $item
    Write-Host "$string"
}

$answ = [System.Windows.MessageBox]::Show("Do you want to proceed?",'PROCEED','YesNo','Info')
if ($answ -eq "No") {    
    Exit
}

# backup job block
Write-Host " "
$RoboCopyBlock = {
    param($final_path, $prefix)
    $filename = $final_path -replace ('\\','-')
    if (Test-Path "C:\DATABACKUP_LOGS\ROBOCOPY_$filename.log" -PathType Leaf) {
        Remove-Item  "C:\DATABACKUP_LOGS\ROBOCOPY_$filename.log" -Force
    }
    New-Item -ItemType file "C:\DATABACKUP_LOGS\ROBOCOPY_$filename.log" > $null
    $source = 'C:\' + $final_path
    $dest = $prefix + '\' + $final_path
    $opts = ("/E", "/Z", "/NP", "/W:5", "/R:5", "/V", "/LOG+:C:\DATABACKUP_LOGS\ROBOCOPY_$filename.log")
    $cmd_args = ($source, $dest, $opts)
    robocopy @cmd_args
}

# launch multithreaded backup jobs
$Time = [System.Diagnostics.Stopwatch]::StartNew()
foreach ($folder in $backup_list.Keys) {
    Write-Host -NoNewline -ForegroundColor Cyan "$folder"
    Start-Job $RoboCopyBlock -Name $folder -ArgumentList $folder, $copiasu > $null
    Write-Host -ForegroundColor Green " JOB STARTED"
}

Start-Sleep 2

<# 
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
#>

# Waiting for jobs completed
While (Get-Job -State "Running") {
    Clear-Host
    Write-Host -ForegroundColor Yellow "*** BACKUP ***"
    
    $total_bytes = 0
    $trasferred_bytes = 0
      
    foreach ($folder in $backup_list.Keys) {
        Write-Host -NoNewline "C:\$folder "
        $source_path = 'C:\' + $folder
        $source_size = $backup_list[$folder]
        $dest_path = $copiasu + '\' + $folder
        $output = robocopy $dest_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Byte:\s+(\d+)\s+\d+" > $null
        $dest_size = $Matches[1]
            
        $total_bytes += $source_size
        $trasferred_bytes += $dest_size

        if ($source_size -gt 1KB) {
            $percent = ($dest_size / $source_size)*100
        } else {
            $percent = 100
        }
        if ($percent -lt 100) {
            $formattato = '{0:0.0}' -f $percent
            Write-Host -ForegroundColor Cyan "$formattato%"
        } else {
            Write-Host -ForegroundColor Green "100%"
        }
    }
    
    $percent = ($trasferred_bytes / $total_bytes)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent
    $CurrentTime = $Time.Elapsed
    $estimated = [int]((($CurrentTime.TotalSeconds/$percent) * (100 - $percent)) / 60)

    <# 
    $label.Text = "Progress: $formattato% - $estimated mins to end"
    if ($progress -ge 100) {
        $bar.Value = 100
    } else {
        $bar.Value = $progress
    }

    # refreshing the progress bar
    [System.Windows.Forms.Application]::DoEvents()
    #>

    Write-Host " "
    Write-Host -ForegroundColor Yellow  "TOTAL PROGRESS: $formattato% - $estimated mins to end"
    Start-Sleep 5
}

# $form_bar.Close()

$joblog = Get-Job | Receive-Job # get job output
Remove-Job * # Cleanup

# Size check
Write-Host " "
Write-Host -NoNewline "Size check..."
foreach ($folder in $backup_list.Keys) {
    $source = 'C:\' + $folder
    $source_size = $backup_list[$folder]
    $dest = $copiasu + '\' + $folder
    $output = robocopy $dest c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
    $output = [system.String]::Join(" ", $output)
    $output -match "Byte:\s+(\d+)\s+\d+" > $null
    $dest_size = $Matches[1]

    $foldername = $folder -replace ('\\','-')                      
    if ($dest_size -lt $source_size) { # backup job failed
        Clear-Host
        $diff = $source_size - $dest_size
        Write-Host "PATH.........: $folder`nSOURCE SIZE..: $source_size bytes`nDEST SIZE....: $dest_size bytes`nDIFF SIZE....: $diff bytes"
    
        $whatif = [System.Windows.MessageBox]::Show("Copy of $folder failed.`nRelaunch backup job?",'ERROR','YesNo','Error')
        if ($whatif -eq "Yes") {
            $opts = ("/E", "/Z", "/NP", "/W:5")
            $cmd_args = ($source, $dest, $opts)    
            Write-Host -ForegroundColor Yellow "RETRY: copy of $folder in progress..."
            Start-Sleep 3
            robocopy @cmd_args
            $whatif = [System.Windows.MessageBox]::Show("Backup of $folder is ok?",'CONFIRM','YesNo','Info')                
            if ($whatif -eq "No") {
                [System.Windows.MessageBox]::Show("Backup $folder manually",'CONFIRM','Ok','Info') > $null
            }
        }
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
    Write-Host -NoNewline "Copying files in C:\Users\$usr..."
    $userfiles = Get-ChildItem "C:\Users\$usr" -Attributes A
    foreach ($afile in $userfiles) {
        Copy-Item "C:\Users\$usr\$afile" -Destination "$copiasu\Users\$usr" -Force > $null
    }
    Write-Host -ForegroundColor Green " DONE"
}

# cleaning temporary
$answ = [System.Windows.MessageBox]::Show("Backup finished. Delete log files?",'END','YesNo','Info')
if ($answ -eq "Yes") {
    Remove-Item "C:\DATABACKUP_LOGS" -Recurse -Force
}
