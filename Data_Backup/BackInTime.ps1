<#
Name......: BackInTime.ps1
Version...: 21.10.1
Author....: Dario CORRADA

This script will restore a backup performed with Data_Backup.ps1
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Data_Backup\\BackInTime\.ps1$" > $null
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
$tmppath = 'C:\BACKINTIME_LOGS'
if (Test-Path $tmppath) {
    Remove-Item "$tmppath" -Recurse -Force
    Start-Sleep 2
}
New-Item -ItemType directory -Path $tmppath > $null

# destination path
$copiasu = 'C:\'

# select source path
$AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
$Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
$OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
$OpenFileDialog.AddExtension = $false
$OpenFileDialog.CheckFileExists = $false
$OpenFileDialog.DereferenceLinks = $true
$OpenFileDialog.Filter = "Folders|`n"
$OpenFileDialog.Multiselect = $false
$OpenFileDialog.Title = "Select source path"
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
$copiada = $OpenFileDialog.FileName

# warn login on remote path (for shared folders)
if ($copiada -match "^\\\\") {
    [System.Windows.MessageBox]::Show("Access on $copiada with your credentials, then click Ok to continue",'WARNING','Ok','Warning')
}

if (!($copiada -match "\\$")) {
    $copiada = $copiada + '\'
}

# collecting path to restore
Write-Host -ForegroundColor Yellow "Checking paths to restore"
$pathlist = @{}
$userslist = @()

$rootpaths = Get-ChildItem -Path $copiada -Attributes D
foreach ($item in $rootpaths) {
    if (!($item.Name -eq 'Users')) {
        Write-Host -NoNewline -ForegroundColor Blue "Processing [$item]..."
        $fullpath = $copiada + $item.Name
        $CommonRobocopyParams = '/E /XJ /R:0 /W:1 /MT:64 /NP /NDL /NC /BYTES /NJH /NJS'
        $StagingLogPath = $tmppath + '\test.log'
        New-Item -Path $StagingLogPath -ItemType File > $null
        $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $fullpath, $StagingLogPath, $CommonRobocopyParams
        Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
        $StagingContent = Get-Content -Path $StagingLogPath
        $TotalFileCount = $StagingContent.Count
        $pathlist[$fullpath] = $TotalFileCount
        Remove-Item $StagingLogPath -Force
        Write-Host -ForegroundColor Green ' DONE'
    }
}

# select users' folders
$userspath = $copiada + 'Users\'
$candidates = Get-ChildItem -Path $userspath -Attributes D
foreach ($item in $candidates) {
    $checkexistingusr = $copiasu + 'Users\' + $item.Name
    $proceed = 'Yes'
    if (!(Test-Path $checkexistingusr)) {
        $proceed = [System.Windows.MessageBox]::Show("User [$item] not found on destination.`nDo you want to proceed?",'PROCEED','YesNo','Warning')
    }
    if ($proceed -eq 'Yes') {
        $userslist += $item.Name
        $candidatepath = $userspath + $item.Name + '\'
        $path2select = Get-ChildItem -Path $candidatepath -Attributes D
        $hsize = 150 + (25 * $path2select.Count)
        $form_panel = FormBase -w 300 -h $hsize -text "$item's folders"
        Label -form $form_panel -x 10 -y 20 -text 'Select folders to restore:' | Out-Null
        $vpos = 45
        $boxes = @()
        foreach ($elem in $path2select) {
            $boxes += CheckBox -form $form_panel -checked $false -x 20 -y $vpos -text $elem.Name
            $vpos += 25
        }
        $vpos += 20
        OKButton -form $form_panel -x 90 -y $vpos -text "Ok" | Out-Null 
        $result = $form_panel.ShowDialog()

        $path2process = @()
        foreach ($box in $boxes) {
            if ($box.Checked -eq $true) {
                $pathname = $box.Text
                $path2process += "$pathname"
            }
        }
        foreach ($subitem in $path2process) {
            Write-Host -NoNewline -ForegroundColor Blue "Processing [$subitem]..."
            $fullpath = $candidatepath + $subitem
            $CommonRobocopyParams = '/E /XJ /R:0 /W:1 /MT:64 /NP /NDL /NC /BYTES /NJH /NJS'
            $StagingLogPath = $tmppath + '\test.log'
            New-Item -Path $StagingLogPath -ItemType File > $null
            $StagingArgumentList = '"{0}" c:\fakepath /LOG:"{1}" /L {2}' -f $fullpath, $StagingLogPath, $CommonRobocopyParams
            Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList
            $StagingContent = Get-Content -Path $StagingLogPath
            $TotalFileCount = $StagingContent.Count
            $pathlist[$fullpath] = $TotalFileCount
            Remove-Item $StagingLogPath -Force
            Write-Host -ForegroundColor Green ' DONE'
        }
    }
}

# summary
Clear-Host
Write-Host -ForegroundColor Yellow "*** PATHS THAT WILL BE RESTORED ***"
$TotalFileToBackup = 0
foreach ($item in ($pathlist.Keys | Sort-Object)) {
    $amount = $pathlist[$item]
    Write-Host -NoNewline -ForegroundColor Cyan "[$item]"
    Write-Host " $amount file(s)"
    $TotalFileToBackup += $amount
}
$answ = [System.Windows.MessageBox]::Show("Do you want to proceed?",'GO','YesNo','Info')
if ($answ -eq "No") {
    Exit
}

# restore job block
$RoboCopyBlock = {
    param($complete_path, $prefixfrom, $prefixto, $logpath)
    $string = $prefixfrom -replace ('\\','\\')
    $complete_path -match "^$string(.+)$" > $null
    $suffix = $matches[1]
    $filename = $suffix -replace ('\\','-')
    if (Test-Path "$logpath\ROBOCOPY_$filename.log" -PathType Leaf) {
        Remove-Item  "$logpath\ROBOCOPY_$filename.log" -Force
    }
    New-Item -ItemType file "$logpath\ROBOCOPY_$filename.log" > $null
    $pathfrom = $complete_path
    $pathto = $prefixto + $suffix

    # for the options see https://superuser.com/questions/814102/robocopy-command-to-do-an-incremental-backup
    $opts = ("/E", "/XJ", "/R:5", "/W:10", "/NP", "/NDL", "/NC", "/NJH", "/ZB", "/MIR", "/LOG+:$logpath\ROBOCOPY_$filename.log")
    if ($prefixfrom -match "^\\\\") {
        $opts += '/COMPRESS'
    }  
    $cmd_args = ($pathfrom, $pathto, $opts)
    robocopy @cmd_args
}

# launch multithreaded restore jobs
foreach ($fullsourcepath in $pathlist.Keys) {
    $string = $copiada -replace ('\\','\\')
    $complete_path -match "^$string(.+)$" > $null
    $suffiz = $matches[1]
    Start-Job $RoboCopyBlock -Name $suffiz -ArgumentList $fullsourcepath, $copiada, $copiasu, $tmppath > $null
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
    Write-Host -ForegroundColor Blue "*** RESTORE ***"
    
    $directoryInfo = Get-ChildItem $tmppath | Measure-Object
    if ($directoryInfo.count -eq 0) {
        Write-Host -ForegroundColor Cyan "Waiting jobs to start..."
    } else {
        $ActiveJobs = 0
        $TotalFilesCopied = 0
        foreach ($item in ($pathlist.Keys | Sort-Object)) {
            $full_path = $item
            $string = $copiada -replace ('\\','\\')
            $full_path -match "^$string(.+)$" > $null
            $suffixo = $matches[1]
            $string = $suffixo -replace ('\\','-')
            $logfile = "$tmppath\ROBOCOPY_$string.log"
            $FilesCopied = 0
            $amount = $pathlist[$item]
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
                    Write-Host -NoNewline "[$suffixo] "
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
Write-Host -ForegroundColor Yellow "COPY CHECK"
foreach ($folder in $pathlist.Keys) {
    $full_path = $folder
    $string = $copiada -replace ('\\','\\')
    $full_path -match "^$string(.+)$" > $null
    $folder = $matches[1]
    $source = $copiada + $folder
    $dest = $copiasu + $folder
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

$answ = [System.Windows.MessageBox]::Show("Remember to grant privileges for other users' paths than yours",'GRANT','Ok','Warning')
<#
*** COMMENTED OUT ***
takeown command seems doesn't work with local account, needs further checks

# assign users' folder attributes
Write-Host " "
Write-Host -ForegroundColor Yellow "SETTING ATTRIBUTES"
foreach ($destuser in $userslist) {
    Write-Host -ForegroundColor Blue -NoNewline "[$destuser] files..."
    foreach ($item in ($pathlist.Keys | Sort-Object)) {
        $full_path = $item
        $string = $copiada -replace ('\\','\\')
        $full_path -match "^$string(.+)$" > $null
        $item = $matches[1]
        $targetfolder = $copiasu + $item
        if ($targetfolder -match "\\Users\\$destuser\\") {
            $logfile = $item -replace ('\\','-')
            $logfile = $tmppath + '\TAKEOWN_' + $logfile + '.log'
            $opts = ("/S", $env:COMPUTERNAME, "/U", $destuser, "/F", $targetfolder, "/R")
            takeown @opts | Out-File $logfile -Encoding ASCII -Append 
        }
    }
    Write-Host -ForegroundColor Green " DONE"
}
#>

# cleaning temporary
$answ = [System.Windows.MessageBox]::Show("Restore finished. Delete log files?",'END','YesNo','Info')
if ($answ -eq "Yes") {
    Remove-Item "$tmppath" -Recurse -Force
}
