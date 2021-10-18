<#
Name......: Automigration.ps1
Version...: 21.10.1
Author....: Dario CORRADA

This script performs the migration of a local user profile from a remote machine
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Automigration\\Automigration\.ps1$" > $null
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

# temporary folder
$tmppath = 'C:\AUTOMIGRATION'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
    New-Item -ItemType file "$tmppath\Automigration.log" > $null
}

# checking source computer
$form_IP = FormBase -w 400 -h 200 -text "SOURCE PC"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(350,30)
$label.Text = "Insert IP address of source machine:"
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
if (Test-Path $prefix) {
    Write-Host -ForegroundColor Green "Source machine connected"
} else {
    [System.Windows.MessageBox]::Show("$prefix not available",'ERROR','Ok','Error') > $null
    Exit
}

# checking users to be migrated
$usrlist = Get-Content -Path "$prefix\AUTOMIGRATION\usrlist.log"
$usr2migrate = @()
foreach ($usr in $usrlist) {
    Write-Host -ForegroundColor Yellow -NoNewline "Checking [$usr] profile..."
    if (Get-LocalUser | Where-Object {$_.Name -eq $usr}) {
        if (net localgroup administrators | Where {$_ -eq $usr}) {
            $usr2migrate += $usr
            Write-Host -ForegroundColor Green " DONE"
        } else {
            $answ = [System.Windows.MessageBox]::Show("User [$usr] is not local admin. Grant privileges?",'USER PROFILE','YesNo','Warning')
            if ($answ -eq "Yes") {
                $ErrorActionPreference = 'Stop'
                try {
                    Add-LocalGroupMember -Group "Administrators" -Member $usr
                    $usr2migrate += $usr
                    Write-Host -ForegroundColor Red " DONE"
                }
                catch {
                    Write-Host -ForegroundColor Red " IGNORED"
                    Write-Host -ForegroundColor Red "$($error[0].ToString())"
                }
                $ErrorActionPreference = 'Inquire'
            } elseif ($answ -eq "No") {
                Write-Host -ForegroundColor Red " IGNORED"
            }
        }
    } else {
        $answ = [System.Windows.MessageBox]::Show("User [$usr] not found. Proceed?",'USER PROFILE','YesNo','Warning')
        if ($answ -eq "Yes") {
            Write-Host -ForegroundColor Red " IGNORED"
        } elseif ($answ -eq "No") {
            Exit
        }
    }
}
Clear-Host
Write-Host -ForegroundColor Yellow "The following users are able to be migrated (granted local admin privileges)"
foreach ($item in $usr2migrate) {
    Write-Host -ForegroundColor Green "[$item]"
}
Pause

# control panel
$form_panel = FormBase -w 400 -h 300 -text "STEPS"
$config_sw = CheckBox -form $form_panel -checked $false -x 20 -y 20 -text "Install standard software"
$migrazione_dati = CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "Data migration"
$network_map = CheckBox -form $form_panel -checked $false -x 20 -y 80 -text "Map shared paths"
$BitLocker_services = CheckBox -form $form_panel -checked $false -x 20 -y 110 -text "Enable BitLocker"
$wifi_profiles = CheckBox -form $form_panel -checked $false -x 20 -y 140 -text "WiFi profiles migration"
OKButton -form $form_panel -x 100 -y 190 -text "Ok"
$result = $form_panel.ShowDialog()

#lancio i singoli moduli selezionati dal pannello di controllo
if ($config_sw.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Installing standard software ***"

    Write-Host -NoNewline "Download software..."
    $download = New-Object net.webclient
    $download.Downloadfile("http://dl.google.com/chrome/install/375.126/chrome_installer.exe", "$tmppath\ChromeSetup.exe")
    $download.Downloadfile("http://ardownload.adobe.com/pub/adobe/reader/win/AcrobatDC/1900820071/AcroRdrDC1900820071_it_IT.exe", "$tmppath\AcroReadDC.exe")
    $download.Downloadfile("https://www.7-zip.org/a/7z1900-x64.exe", "$tmppath\7Zip.exe")
    $download.Downloadfile("https://go.skype.com/windows.desktop.download", "$tmppath\Skype.exe")
    $download.Downloadfile("https://download.ccleaner.com/spsetup132.exe", "$tmppath\Speccy.exe")
    $download.Downloadfile("https://www.revouninstaller.com/download-freeware-version.php", "$tmppath\Revo.exe")
    $download.Downloadfile("https://www.supremocontrol.com/download.aspx?file=Supremo.exe&id_sw=7&ws=supremocontrol.com", "$env:PUBLIC\Desktop\Supremo.exe")
    Write-Host -ForegroundColor Green " DONE"

    Write-Host -NoNewline "Install software..."
    Start-Process -FilePath "$tmppath\ChromeSetup.exe" -Wait
    Start-Process -FilePath "$tmppath\AcroReadDC.exe" -Wait
    Start-Process -FilePath "$tmppath\7Zip.exe" -Wait
    Start-Process -FilePath "$tmppath\Speccy.exe" -Wait
    Start-Process -FilePath "$tmppath\Revo.exe" -Wait
    $answ = [System.Windows.MessageBox]::Show("After installation close all Skype instances to proceed...",'WARNING','Ok','Warning')
    Start-Process -FilePath "$tmppath\Skype.exe" -Wait
    Write-Host -ForegroundColor Green " DONE"
}

if ($migrazione_dati.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Data migration ***"

    $targetpaths = Get-Content -Path "$prefix\AUTOMIGRATION\targetpaths.log"
    $paths2migrate = @()
    foreach ($path in $targetpaths) {
        foreach ($usr in $usr2migrate) {
            if ($path -match "^Users\\$usr\\") {
                $paths2migrate += $path               
            }
        }
    }
    Write-Host "The following paths will be migrated"
    foreach ($item in $paths2migrate) {
        Write-Host -ForegroundColor Cyan "[$item]"
    }
    Pause

    $copiasu = 'C:\'
    $backup_list = @{} # variable in which will added paths to backup
    $root_path = $prefix + '\'

    Write-Host -NoNewline "`n`nChecking paths..."
    foreach ($item in $paths2migrate) {
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
        } else {
            Write-Host -NoNewline "`n`nPath [$full_path] not found "
            Write-Host -ForegroundColor Red "IGNORED"
            Pause
            Write-Host -NoNewline "`nChecking paths..."
        }
    }
    Write-Host -ForegroundColor Green " DONE"

    # backup job block
    Write-Host " "
    $RoboCopyBlock = {
        param($final_path, $pathat, $logpath)
        $filename = $final_path -replace ('\\','-')
        if (Test-Path "$logpath\ROBOCOPY_$filename.log" -PathType Leaf) {
            Remove-Item  "$logpath\ROBOCOPY_$filename.log" -Force
        }
        New-Item -ItemType file "$logpath\ROBOCOPY_$filename.log" > $null
        $source = 'C:\' + $final_path
        $dest = $pathat + '\' + $final_path

        # for the options see https://superuser.com/questions/814102/robocopy-command-to-do-an-incremental-backup
        $opts = ("/E", "/XJ", "/R:5", "/W:10", "/NP", "/NDL", "/NC", "/NJH", "/ZB", "/MIR", "/LOG+:$logpath\ROBOCOPY_$filename.log")
        if ($pathat -match "^\\\\") {
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

    # file backup from Users folder
    foreach ($usr in $usr2migrate) {
        $prefisso = $prefix + '\Users\' + $usr
        Write-Host -NoNewline "Copying individual user profile files for [$usr]..."
        $userfiles = Get-ChildItem "$prefisso" -Attributes A
        foreach ($afile in $userfiles) {
            Copy-Item "$prefisso\$afile" -Destination "C:\Users\$usr" -Force > $null
        }
        Write-Host -ForegroundColor Green " DONE"
    }
}

if ($network_map.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Mapping shared paths ***"

    $net_path = $prefix + "\AUTOMIGRATION\NetworkDrives.log"
    if (Test-Path $net_path -PathType Leaf) {
        [string[]]$net_list = Get-Content -Path $net_path
        $Network = New-Object -ComObject "Wscript.Network"
        foreach ($newnet in $net_list) {
            $letter,$fullpath = $newnet.split(';')
            $Network.MapNetworkDrive($letter, $fullpath, 1)
            Write-Host -ForegroundColor Green $fullpath
        }
    } else {
        Write-Host -ForegroundColor Red "No path found"
    }
}

if ($BitLocker_services.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Enabling BitLocker ***"

    $form_modalita = FormBase -w 300 -h 175 -text "CRYPTO LEVEL"
    $tpm = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "TPM"
    $pin  = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "TPM + PIN"
    OKButton -form $form_modalita -x 90 -y 90 -text "Ok"
    $result = $form_modalita.ShowDialog()
    if ($result -eq "OK") {
        if ($tpm.Checked) {
            $outarget = "null"
        } elseif ($pin.Checked) {
            $form = FormBase -w 520 -h 200 -text "PIN"
            $font = New-Object System.Drawing.Font("Arial", 12)
            $form.Font = $font
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10,20)
            $label.Size = New-Object System.Drawing.Size(500,30)
            $label.Text = "6-digit PIN:"
            $form.Controls.Add($label)                  
            $pwds = New-Object System.Windows.Forms.TextBox
            $pwds.Location = New-Object System.Drawing.Point(10,60)
            $pwds.Size = New-Object System.Drawing.Size(450,30)
            $pwds.PasswordChar = '*'
            $form.Controls.Add($pwds)
            $OKButton = New-Object System.Windows.Forms.Button
            OKButton -form $form -x 200 -y 120 -text "Ok"
            $form.Topmost = $true
            $result = $form.ShowDialog()
            $thepin = $pwds.Text
        }    
    }

    $ErrorActionPreference = 'Stop'
    try {
        # https://docs.microsoft.com/en-us/powershell/module/bitlocker/enable-bitlocker?view=windowsserver2019-ps
        if ($pin.Checked) {
            $SecureString = ConvertTo-SecureString $thepin -AsPlainText -Force
            Enable-BitLocker -MountPoint "C:" -EncryptionMethod Aes256 -UsedSpaceOnly -Pin $SecureString -TPMandPinProtector
        } else {
            Enable-BitLocker -MountPoint "C:" -EncryptionMethod Aes256 -UsedSpaceOnly -TpmProtector
        }
        Add-BitLockerKeyProtector -MountPoint "C:" -RecoveryPasswordProtector
        Write-Host "Wait..."
        Start-Sleep 5
    }
    catch {
        Write-Host -ForegroundColor Red "FAILED"
        Write-Host -ForegroundColor Red "$($error[0].ToString())"
        Pause
    }
    $ErrorActionPreference = 'Inquire'

    # reminder
    if ($pin.Checked) {
        Write-Host -NoNewline "`nThe BitLocker PIN you have setted is "
        Write-Host -ForegroundColor Cyan "$thepin"
    }
    $BLkey = Get-BitLockerVolume
    foreach ($item in $BLkey) {
        if ($item.MountPoint -match 'C') {
            $RecoveryKey = [string]($item.KeyProtector).RecoveryPassword
            Write-Host -NoNewline "`nThe Recovery Key is "
            Write-Host -ForegroundColor Cyan "$RecoveryKey"    
        }
    }
    Pause
}

if ($wifi_profiles.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Migrating WiFi profiles ***"

    $wifi_profile_list = Get-ChildItem "$prefix\AUTOMIGRATION\wifi_profiles"
    foreach ($wifi_profile in $wifi_profile_list) {
        netsh wlan add profile filename="$prefix\AUTOMIGRATION\wifi_profiles\$wifi_profile" user=current
    }
}

# cleaning temporary
$answ = [System.Windows.MessageBox]::Show("Data migration accomplished. Delete log files?",'END','YesNo','Info')
if ($answ -eq "Yes") {
    Remove-Item "$tmppath" -Recurse -Force
}

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}
