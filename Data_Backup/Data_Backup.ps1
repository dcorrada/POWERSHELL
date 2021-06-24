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
$tmppath = 'C:\TEMPSOFTWARE'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
}

$copiasu = Read-Host "Insert destination path"

# login on remote path (for shared folders)
if ($copiasu -match "^\\\\") {
    [System.Windows.MessageBox]::Show("Access on $copiasu with your credentials, then click Ok to continue",'WARNING','Ok','Warning')
}

# check destination path
if (Test-Path $copiasu) {
    $copiasu = $copiasu + "\$env:USERNAME"
    New-Item -ItemType directory -Path $copiasu > $null
} else {
    [System.Windows.MessageBox]::Show("$copiasu not found",'ERROR','Ok','Error') > $null
    Exit
}

# copying Chrome bookmarks
$bookmarks = "C:\Users\$env:USERNAME\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"
if (Test-Path $bookmarks -PathType Leaf) {
    Write-Host -NoNewline "Copying Chrome bookmarks..."
    New-Item -ItemType directory -Path "$copiasu\Users\$env:USERNAME\AppData\Local\Google\Chrome\User Data\Default" > $null
    Copy-Item $bookmarks -Destination "$copiasu\Users\$env:USERNAME\AppData\Local\Google\Chrome\User Data\Default" > $null
    Write-Host -ForegroundColor Green " DONE"
}

# copying Outlook layout
$outlook_aspect = "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook\Outlook.xml"
if (Test-Path $outlook_aspect -PathType Leaf) {
    Write-Host -NoNewline "Copying Outlook layout..."
    New-Item -ItemType directory -Path "$copiasu\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook" > $null
    Copy-Item $outlook_aspect -Destination "$copiasu\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook" > $null
    Write-Host -ForegroundColor Green " DONE"
}

Write-Host -NoNewline "Searching paths to backup..."
$backup_list = @{} # variable in which I will add paths to backup

# backup paths list; empty folders are excluded from backup    
[string[]]$allow_list = Get-Content -Path "$repopath\Backup_Dati\allow_list.log"
$allow_list = $allow_list -replace ('\$username', $env:USERNAME)             
foreach ($folder in $allow_list) {
    $full_path = 'C:\' + $folder
    if (Test-Path $full_path) {
        $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Byte:\s+(\d+)\s+\d+" > $null
        $size = $Matches[1]
        if ($size -gt 1KB) {
            $backup_list[$folder] = $size
            Write-Host -NoNewline "."
        }
    }
}

# paths excluded from backup;
[string[]]$exclude_list = Get-Content -Path "$repopath\Backup_Dati\exclude_list.log"
$root_path = 'C:\'
$remote_root_list = Get-ChildItem $root_path -Attributes D
$elenco = @();
foreach ($folder in $remote_root_list.Name) {
    if (!($exclude_list -contains $folder)) {
        $elenco += $folder
    }
}
$string = [system.String]::Join("`r`n", $elenco)
$form_folders = FormBase -w 400 -h 275 -text "FOLDER LIST IN C:\"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(350,30)
$label.Text = "Delete folders you don't want to backup:"
$form_folders.Controls.Add($label)
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Multiline = $true
$textBox.Scrollbars = "Vertical"
$textBox.Location = New-Object System.Drawing.Point(10,50)
$textBox.Size = New-Object System.Drawing.Size(350,100)
$textBox.Text = $string
$form_folders.Controls.Add($textBox)
OKButton -form $form_folders -x 100 -y 175 -text "Ok"
$form_folders.Add_Shown({$textBox.Select()})
$result = $form_folders.ShowDialog()
$elenco = $textBox.Text.Split("`r`n")   
foreach ($folder in $elenco) {
    if (!($folder -eq "")) {
        if (!($exclude_list -contains $folder)) {
            $full_path = $root_path + $folder
            $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
            $output = [system.String]::Join(" ", $output)
            $output -match "Byte:\s+(\d+)\s+\d+" > $null
            $size = $Matches[1]
            $backup_list[$folder] = $size
            Write-Host -NoNewline "."
        }
    }
}

# folders list to backup from C:\Users\username
$exclude_list = (
    "AppData",
    "Links",
    "OneDrive",
    "Saved Games",
    "Searches",
    "3D Objects",
    ".cisco",
    ".config",
    "Dropbox"
)
$root_path = "C:\Users\$env:USERNAME\"
$remote_root_list = Get-ChildItem $root_path -Attributes D
foreach ($folder in $remote_root_list.Name) {
    if (!($exclude_list -contains $folder)) {
        $full_path = $root_path + $folder
        $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Byte:\s+(\d+)\s+\d+" > $null
        $size = $Matches[1]
        $full_path = "Users\$env:USERNAME\$folder"
        $backup_list[$full_path] = $size
        Write-Host -NoNewline "."
    }
}

Write-Host -ForegroundColor Green " DONE"

# backup job block
$RoboCopyBlock = {
    param($final_path, $prefix)
    $filename = $final_path.Replace('\','-')
    if (Test-Path "C:\TEMPSOFTWARE\ROBOCOPY_$filename.log" -PathType Leaf) {
        Remove-Item  "C:\TEMPSOFTWARE\ROBOCOPY_$filename.log" -Force
    }
    New-Item -ItemType file "C:\TEMPSOFTWARE\ROBOCOPY_$filename.log" > $null
    $source = 'C:\' + $final_path
    $dest = $prefix + '\' + $final_path
    $opts = ("/E", "/Z", "/NP", "/W:5", "/R:5", "/V", "/LOG+:C:\TEMPSOFTWARE\ROBOCOPY_$filename.log")
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

Start-Sleep 10

# progress bar
$form_bar = New-Object System.Windows.Forms.Form
$form_bar.Text = "TRANSFER RATE"
$form_bar.Size = New-Object System.Drawing.Size(600,200)
$form_bar.StartPosition = "manual"
$form_bar.Location = '1320,840'
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
$form_bar.Topmost = $true
$form_bar.Show() | out-null
$form_bar.Focus() | out-null

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
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent
    $CurrentTime = $Time.Elapsed
    $estimated = [int]((($CurrentTime.TotalSeconds/$percent) * (100 - $percent)) / 60)
    $label.Text = "Progress: $formattato% - $estimated mins to end"
    if ($progress -ge 100) {
        $bar.Value = 100
    } else {
        $bar.Value = $progress
    }
    $form_bar.Refresh()
    Write-Host " "
    Start-Sleep 5
}

$form_bar.Close()

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

    $foldername = $folder.Replace('\','-')                      
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
Write-Host -NoNewline "Copying files in C:\Users\$env:USERNAME..."
$userfiles = Get-ChildItem "C:\Users\$env:USERNAME" -Attributes A
foreach ($afile in $userfiles) {
    Copy-Item "C:\Users\$env:USERNAME\$afile" -Destination "$copiasu\Users\$env:USERNAME" -Force > $null
}
Write-Host -ForegroundColor Green " DONE"

# cleaning temporary
$answ = [System.Windows.MessageBox]::Show("Backup finished. Delete log files?",'END','YesNo','Info')
if ($answ -eq "Yes") {
    Remove-Item "C:\TEMPSOFTWARE" -Recurse -Force
}
