<#
Name......: Automigration.ps1
Version...: 20.2.1
Author....: Dario CORRADA

This script performs data migration from source computer that have volume C: in smbshare

NOTE: in older versions of robocopy the stdout slughtly differs, causing malfunction of the script. 
In case replace all occurrences of "Byte:\s+(\d+)\s+\d+" to "Bytes :\s+(\d+)\s+\d+"
#>

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

# getting the current directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Automigration\.ps1$" > $null
$workdir = $matches[1]

# elevate script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# temporary working directory
$tmppath = 'C:\TEMPSOFTWARE'
if (!(Test-Path $tmppath)) {
   # New-Item -ItemType directory -Path $tmppath > $null
}

# import module for dialog boxes
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\DataMigration\\Automigration\.ps1$" > $null
$modpath = $matches[1] + "\Modules"
Import-Module -Name "$modpath\Forms.psm1"

# temporary directory
$tmppath = 'C:\TEMPSOFTWARE'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
    New-Item -ItemType file "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" > $null
}

# searching source PC
$form_IP = FormBase -w 400 -h 200 -text "SOURCE PC"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(350,30)
$label.Text = "Insert local IP or hostname of the source PC:"
$form_IP.Controls.Add($label)
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(10,60)
$textBox.Size = New-Object System.Drawing.Size(350,30)
$form_IP.Controls.Add($textBox)
OKButton -form $form_IP -x 100 -y 100 -text "Ok"
$form_IP.Add_Shown({$textBox.Select()})
$result = $form_IP.ShowDialog()
if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
    $ipaddress = $textBox.Text
    $prefix = '\\' + $ipaddress + '\C'
}

# checking Office version
# needs patches for previous Office versions
Write-Host -NoNewline "`n`nChecking Office version... "
$record = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
          Get-ItemProperty | Where-Object {$_.DisplayName -match 'Office' } | Select-Object -Property DisplayName, UninstallString
$OfficeVer = '15'
foreach ($elem in $record) {
    if ($elem -match 'Office 16') {
        $OfficeVer = '16'
    }
}
if ($OfficeVer -eq '16') {
    Write-Host -ForegroundColor Cyan "Office 365 installed"
} else {
    Write-Host -ForegroundColor Cyan "Office 2013 installed"
}

# control panel
$form_panel = FormBase -w 320 -h 350 -text "STEPS"
$config_sw = CheckBox -form $form_panel -checked $true -x 20 -y 20 -text "Configuring Office"
$config_sw.Checked = $false
$migrazione_dati = CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "Data migration"
$network_map = CheckBox -form $form_panel -checked $true -x 20 -y 80 -text "Mapping network shortcuts"
$setting_taskbar = CheckBox -form $form_panel -checked $true -x 20 -y 110 -text "Setting taskbar"
$setting_taskbar.Checked = $false
$machine_certificate = CheckBox -form $form_panel -checked $true -x 20 -y 140 -text "Rquest user certificates"
$machine_certificate.Enabled = $false
$machine_certificate.Checked = $false
$wifi_profiles = CheckBox -form $form_panel -checked $true -x 20 -y 170 -text "Import WiFi profiles"
$sendlog = CheckBox -form $form_panel -checked $true -x 20 -y 200 -text "Sending log mail"
$sendlog.Enabled = $false
$sendlog.Checked = $false
OKButton -form $form_panel -x 100 -y 250 -text "Ok"
$result = $form_panel.ShowDialog()

# checking connection
[System.Windows.MessageBox]::Show("From File Explorer access once to $prefix`nThen click Ok to continue",'DISCLAIMER','Ok','Info') > $null
if (!(Test-Path $prefix)) {
    [System.Windows.MessageBox]::Show("Unable to reach $prefix",'ERROR','Ok','Error') > $null
    Exit
}

# retrieve logs from source PC
$log_pcvecchio = $prefix + "\TEMPSOFTWARE\LocalAdmin-Cshared.log"
Copy-Item $log_pcvecchio -Destination "C:\TEMPSOFTWARE" > $null
$logcontent = Get-Content "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log"
$pcname = $logcontent[3]
"*** DATA MIGRATION FROM: $pcname ***" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append

# configuring Office
if ($config_sw.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Configuring Office ***"

    # Autoconfig silent of Outlook
    Write-Host -NoNewline "Configuring Outlook..."
    if ($OfficeVer -eq '15') {
        $regPath = 'HKCU:\Software\Microsoft\Office\15.0\Outlook\AutoDiscover'
    } elseif ($OfficeVer -eq '16') {
        $regPath = 'HKCU:\Software\Microsoft\Office\16.0\Outlook\AutoDiscover'
    }
    New-Item $regPath -Force | Out-Null
    New-ItemProperty $regPath -Name ZeroConfigExchange -Value 1 -Force  | Out-Null    
    Start-Sleep 5
    Start-Process outlook
    Start-Sleep 20
    Write-Host -ForegroundColor Green " DONE"

    # Autoconfig silent of Office LanguagePack
    [System.Windows.MessageBox]::Show("Setting Office language [it-IT]...",'CONFIGURAZIONE SW','Ok','Info') > $null
    Write-Host -NoNewline "Settaggio lingua Office..."

    if ($OfficeVer -eq '15') {
        $regPath = 'HKCU:\Software\Microsoft\Office\15.0\Common\LanguageResources'
        Set-ItemProperty $regPath -Name UILanguage -Value 1040 -Force
        Set-ItemProperty $regPath -Name HelpLanguage -Value 1040 -Force
        Set-ItemProperty $regPath -Name UIFallback -Value {1040;0;1033} -Force
        Set-ItemProperty $regPath -Name HelpFallback -Value {1040;0;1033} -Force
        New-ItemProperty $regPath -Name FollowSystemUI -Value Off -Force  | Out-Null
    } elseif ($OfficeVer -eq '16') {
        $regPath = 'HKCU:\Software\Microsoft\Office\16.0\Common\LanguageResources'
        Set-ItemProperty $regPath -Name UILanguageTag -Value "it-IT" -Force
        Set-ItemProperty $regPath -Name HelpLanguageTag -Value "it-IT" -Force
        Set-ItemProperty $regPath -Name UIFallbackLanguages -Value "it-it;x-none;en-us" -Force
        Set-ItemProperty $regPath -Name HelpFallbackLanguages -Value "it-it;x-none;en-us" -Force
        New-ItemProperty $regPath -Name FollowSystemUI -Value Off -Force  | Out-Null
    }
    Write-Host -ForegroundColor Green " DONE"
}

# performing data migration
if ($migrazione_dati.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Data migration ***"
    $ErrorActionPreference= 'SilentlyContinue'

    # retrieving Chrome bookmarks
    if ($install_Chrome.Checked -eq $false) {
        Start-Process -Wait "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    }
    $bookmarks = $prefix + "\Users\$env:USERNAME\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"
    if (Test-Path $bookmarks -PathType Leaf) {
        Write-Host -NoNewline "Copying Chrome bookmarks..."
        Copy-Item $bookmarks -Destination "C:\Users\$env:USERNAME\AppData\Local\Google\Chrome\User Data\Default" > $null
        Write-Host -ForegroundColor Green " DONE"
    }

    # retrieving Outlook layout
    $outlook_aspect = $prefix + "\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook\Outlook.xml"
    if (Test-Path $outlook_aspect -PathType Leaf) {
        Write-Host -NoNewline "Copying Outlook layout..."
        Remove-Item "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook\Outlook.xml" -Force
        Copy-Item $outlook_aspect -Destination "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook" > $null
        Write-Host -ForegroundColor Green " DONE"
    }

    Write-Host -NoNewline "Looking for paths to be migrated..."
    $backup_list = @{} # variable in which the paths will be listed
    
    # retrieving paths to be migrated 
    [string[]]$allow_list = Get-Content -Path "$workdir\allow_list.log"
    $allow_list = $allow_list -replace ('\$username', $env:USERNAME)             
    foreach ($folder in $allow_list) {
        $full_path = $prefix + '\' + $folder
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

    # generating list of paths to be escluded from migration
    [string[]]$exclude_list = Get-Content -Path "$workdir\exclude_list.log"
    $root_path = $prefix + '\'
    $remote_root_list = Get-ChildItem $root_path -Attributes D
    $elenco = @();
    foreach ($folder in $remote_root_list.Name) {
        if (!($exclude_list -contains $folder)) {
            $elenco += $folder
        }
    }
    $string = [system.String]::Join("`r`n", $elenco)
    $form_folders = FormBase -w 400 -h 275 -text "LIST FOLDERS IN C:\"
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(350,30)
    $label.Text = "Delete those folder you don't want to transfer:"
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

    # add paths to be migrated downstream of C:\Users\[username]
    $exclude_list = ( # exclude list may be updated according to user flavours
        "Links",
        "OneDrive",
        "DropBox",
        "Searches",
        ".config"
    )
    $root_path = "$prefix\Users\$env:USERNAME\"
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

    # data migration single job
    $RoboCopyBlock = {
        param($final_path,$prefisso)
        $filename = $final_path.Replace('\','-')
        if (Test-Path "C:\TEMPSOFTWARE\ROBOCOPY_$filename.log" -PathType Leaf) {
            Remove-Item "C:\TEMPSOFTWARE\ROBOCOPY_$filename.log" -Force
        }
        New-Item -ItemType file "C:\TEMPSOFTWARE\ROBOCOPY_$filename.log" > $null
        $source_path = $prefisso + '\' + $final_path
        $dest_path = 'C:\' + $final_path
        $opts = ("/E", "/Z", "/NP", "/IS", "/IT", "/W:5", "/R:5", "/V", "/LOG+:C:\TEMPSOFTWARE\ROBOCOPY_$filename.log")
        $cmd_args = ($source_path, $dest_path, $opts)    
        robocopy @cmd_args
    }

    # launch data migration jobs in parallel
    $Time = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($folder in $backup_list.Keys) {
        Write-Host -NoNewline -ForegroundColor Cyan "$folder"
        Start-Job $RoboCopyBlock -Name $folder -ArgumentList $folder, $prefix > $null
        Write-Host -ForegroundColor Green " JOB STARTED"
    }
    Start-Sleep 10

    # progress bar
    $form_bar = New-Object System.Windows.Forms.Form
    $form_bar.Text = "DATA MIGRATION"
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

    # wait for all jobs finishing
    While (Get-Job -State "Running") {
        Clear-Host
        Write-Host -ForegroundColor Yellow "*** DATA MIGRATION IN PROGRESS ***"

        $total_bytes = 0
        $trasferred_bytes = 0
        
        foreach ($folder in $backup_list.Keys) {
            Write-Host -NoNewline "C:\$folder "
            $source_path = $prefix + '\' + $folder
            $source_size = $backup_list[$folder]
            $dest_path = 'C:\' + $folder
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
        $label.Text = "Total progress: $formattato% - $estimated mins to go"
        $bar.Value = $progress
        $form_bar.Refresh()

        Write-Host " "
        Start-Sleep 5
    }

    $form_bar.Close()

    $joblog = Get-Job | Receive-Job # retrieve output job
    Remove-Job * # Cleanup

    # size check
    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Data migration" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
    Write-Host " "
    Write-Host -NoNewline "Size check..."
    foreach ($folder in $backup_list.Keys) {
        $source_path = $prefix + '\' + $folder
        $source_size = $backup_list[$folder]
        $dest_path = 'C:\' + $folder
        $output = robocopy $dest_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Byte:\s+(\d+)\s+\d+" > $null
        $dest_size = $Matches[1]

        $foldername = $folder.Replace('\','-')                      
        if ($dest_size -ge $source_size) { # copy has been accomplishied
            $logcontent = Get-Content "C:\TEMPSOFTWARE\ROBOCOPY_$foldername.log"
            for ($i = 0; $i -lt $logcontent.Count; $i++) {
                if ($logcontent[$i] -match "^     Dest :") {
                    $logcontent[$i] | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
                    "------------------------------------------------------------------------------" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
                } elseif ($logcontent[$i] -match "^               Total    Copied   Skipped  Mismatch    FAILED    Extras$") {
                    $logcontent[$i] | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
                } elseif ($logcontent[$i] -match "^   Files :\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+$") {
                    $logcontent[$i] | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
                }
            }
            "------------------------------------------------------------------------------" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
        } else { # copy is failed
            [System.Windows.MessageBox]::Show("Ooops! Something goes wrong...`ncheck C:\TEMPSOFTWARE\ROBOCOPY_$foldername.log",'WARNING','Ok','Warning')
                                   
            Clear-Host
            $diff = $source_size - $dest_size
            Write-Host "PATH.........: $folder`nSOURCE SIZE..: $source_size bytes`nDEST SIZE....: $dest_size bytes`nDIFF SIZE....: $diff bytes"

            $whatif = [System.Windows.MessageBox]::Show("Relaunch copy of $folder?",'INFO','YesNo','Error')
            if ($whatif -eq "Yes") {
                New-Item -ItemType file "C:\TEMPSOFTWARE\ROBOCOPY-RETRY_$foldername.log" > $null
                $opts = ("/E", "/Z", "/NP", "/IS", "/IT", "/W:5", "/V", "/TEE", "/LOG+:C:\TEMPSOFTWARE\ROBOCOPY-RETRY_$foldername.log")
                $cmd_args = ($source_path, $dest_path, $opts)    
                Write-Host -ForegroundColor Yellow "RETRY: copy $folder in progress..."
                Start-Sleep 3
                robocopy @cmd_args
                notepad "C:\TEMPSOFTWARE\ROBOCOPY-RETRY_$foldername.log"
                $whatif = [System.Windows.MessageBox]::Show("Copy of $folder is ok?",'CONFIRM','YesNo','Info')                
                if ($whatif -eq "Yes") {
                    $logcontent = Get-Content "C:\TEMPSOFTWARE\ROBOCOPY_$foldername.log"
                    for ($i = 0; $i -lt $logcontent.Count; $i++) {
                        if ($logcontent[$i] -match "^     Dest :") {
                            $logcontent[$i] | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
                            "------------------------------------------------------------------------------" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
                        } elseif ($logcontent[$i] -match "^               Total    Copied   Skipped  Mismatch    FAILED    Extras$") {
                            $logcontent[$i] | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
                        } elseif ($logcontent[$i] -match "^   Files :\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+\s+\d+$") {
                            $logcontent[$i] | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
                        }
                    }
                    "------------------------------------------------------------------------------" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
                } else {
                    [System.Windows.MessageBox]::Show("You should copy $folder manually",'INFO','Ok','Info') > $null
                }
            }
        }    
    }
    $ErrorActionPreference= 'Inquire'
    Write-Host -ForegroundColor Green " DONE"

    # check attributes for hidden folders
    Write-Host " "
    Write-Host -NoNewline "Check attributes..."
    foreach ($folder in $backup_list.Keys) {
        $dest_path = 'C:\' + $folder
        attrib -s -h $dest_path
    }
    Write-Host -ForegroundColor Green " DONE"

    # Copy of individual files in C:\Users\username
    Write-Host -NoNewline "Copia files in C:\Users\$env:USERNAME..."
    $userfiles = Get-ChildItem "C:\Users\$env:USERNAME" -Attributes A
    foreach ($afile in $userfiles) {
        Copy-Item "$prefix\Users\$env:USERNAME\$afile" -Destination "C:\Users\$env:USERNAME" -Force > $null
    }
    Write-Host -ForegroundColor Green " DONE"
}

if ($network_map.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Mapping network shortcuts ***"
    $net_path = $prefix + "\TEMPSOFTWARE\NetworkDrives.log"
    if (Test-Path $net_path -PathType Leaf) {
        Copy-Item $net_path -Destination "C:\TEMPSOFTWARE"
        [string[]]$net_list = Get-Content -Path 'C:\TEMPSOFTWARE\NetworkDrives.log'
        $Network = New-Object -ComObject "Wscript.Network"
        foreach ($newnet in $net_list) {
            $letter,$fullpath = $newnet.split(';')
            $Network.MapNetworkDrive($letter, $fullpath, 1)
            Write-Host -ForegroundColor Green $fullpath
        }
    } else {
        Write-Host -ForegroundColor Cyan "No shortcut found"
    }

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Mapping network shortcuts" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($setting_taskbar.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Setting taskbar ***"
    Remove-Item "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" -Recurse -Force
    Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" -Recurse -Force
    $source_path = $prefix + '\TEMPSOFTWARE\TaskBar'
    $dest_path = 'C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar'
    $opts = ("/E", "/Z", "/NP")
    $cmd_args = ($source_path, $dest_path, $opts)    
    robocopy @cmd_args  > $null
    $regfile = $prefix + "\TEMPSOFTWARE\custom_pinned.reg"
    Copy-Item $regfile -Destination "C:\TEMPSOFTWARE" > $null
    Start-Process "C:\TEMPSOFTWARE\custom_pinned.reg" -Wait
    Write-Host -ForegroundColor Green " DONE"

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Setting taskbar" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($machine_certificate.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Request user certificates ***"

    # list available certificates
    $formlist = FormBase -w 400 -h 200 -text "CERTIFICATES"
    $DropDown = new-object System.Windows.Forms.ComboBox
    $DropDown.Location = new-object System.Drawing.Size(10,60)
    $DropDown.Size = new-object System.Drawing.Size(350,30)
    $font = New-Object System.Drawing.Font("Arial", 12)
    $DropDown.Font = $font
    foreach ($elem in ("Utente", "EFS di base")) { # modify here the typology of certificates
        $DropDown.Items.Add($elem)  > $null
    }
    $formlist.Controls.Add($DropDown)
    $DropDownLabel = new-object System.Windows.Forms.Label
    $DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
    $DropDownLabel.size = new-object System.Drawing.Size(500,30) 
    $DropDownLabel.Text = "Select certificate"
    $formlist.Controls.Add($DropDownLabel)
    OKButton -form $formlist -x 100 -y 100 -text "Ok"
    $formlist.Add_Shown({$DropDown.Select()})
    $result = $formlist.ShowDialog()
    $selected = $DropDown.Text

    if ($selected -match "Utente") {
        $tag = "User"
    } elseif ($metti.Checked) {
        $tag = "EFS"
    }

    # Read documentation of cmdlet certreq:
    # https://docs.microsoft.com/it-it/windows-server/administration/windows-commands/certreq_1  
    $certreq_log = certreq -enroll -user -policyserver * $tag

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Request user certificates" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($wifi_profiles.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Import WiFi profiles ***"

    $wifi_profile_sourcepath = $prefix + "\TEMPSOFTWARE\wifi_profiles"
    Copy-Item -Recurse $wifi_profile_sourcepath -Destination "C:\TEMPSOFTWARE\" > $null
    
    Write-Host "Restore wireless network profile"
    $wifi_profile_list = Get-ChildItem "C:\TEMPSOFTWARE\wifi_profiles"
    $exclude_list = ("Administered.xml") # insert here the list of profile you don't want to add
    foreach ($wifi_profile in $wifi_profile_list) {
        if ($exclude_list -contains $wifi_profile) {
            Write-Host "Skip $wifi_profile"
        } else {
            netsh wlan add profile filename="C:\TEMPSOFTWARE\wifi_profiles\$wifi_profile" user=current
        }
    }
    Remove-Item "C:\TEMPSOFTWARE\wifi_profiles" -Recurse -Force

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Import WiFi profiles" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($sendlog.Checked -eq $true) {
    if (Test-Path "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -PathType Leaf) {
        $header = 'Data migration log' + $textBox.Text
    
        $Outlook = New-Object -ComObject Outlook.Application
        $Mail = $Outlook.CreateItem(0)
        $Mail.To = "dario.corrada@gmail.com" # insert your email address 
        $Mail.Subject = $header
        $Mail.Body = "Data migration performed on $env:COMPUTERNAME"
        $Mail.Attachments.Add("C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log");
        $Mail.Send()
        $Outlook.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
}

# pulizia temporanei
$answ = [System.Windows.MessageBox]::Show("Clear temporary files?",'TEMPUS','YesNo','Info')
if ($answ -eq "Yes") {
    Remove-Item "C:\TEMPSOFTWARE" -Recurse -Force
} else {
    [System.Windows.MessageBox]::Show("Temporary files are stored in C:\TEMPSOFTWARE",'INFO','Ok','Info')
}

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {
    Restart-Computer
}
