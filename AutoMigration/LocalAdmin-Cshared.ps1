<#
Name......: LocalAdmin-Cshared.ps1
Version...: 21.10.1
Author....: Dario CORRADA

This script will set user as local-admin and share the entire C: volume
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Automigration\\LocalAdmin-Cshared\.ps1$" > $null
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
if (Test-Path $tmppath) {
    Remove-Item "$tmppath" -Recurse -Force
    Start-Sleep 2
}
New-Item -ItemType directory -Path $tmppath > $null

# preview of folders to be migrated
Write-Host -ForegroundColor Yellow "Looking for paths to be migrated..."
$backup_list = @{} # list of paths to be migrated

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

# listing paths
[string[]]$exclude_list = Get-Content -Path $excludefile
foreach ($usr in $usrlist) {
    $root_usr = 'C:\Users\' + $usr + '\'
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



Pause




Write-Host "Cerco i percorsi di backup..."


# lista cartelle su cui si dovrebbe fare la migrazione; se la cartella non e' vuota la metto nella lista dei backup    
[string[]]$allow_list = Get-Content -Path "\\itmilitgroup\SD_Utilities\SCRIPT\MigrazioneDati\Automigrazione_allow_list.log"
$allow_list = $allow_list -replace ('\$username', $env:USERNAME)             
foreach ($folder in $allow_list) {
    $full_path = 'C:\' + $folder
    if (Test-Path $full_path) {
        $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Bytes :\s+(\d+)\s+\d+" > $null
        $size = $Matches[1]
        if ($size -gt 1KB) {
            $backup_list[$folder] = $size
        }
    }
}

# lista cartelle da escludere dalla migrazione; se in C:\ ci sono altre cartelle oltre a queste le includo nella lista dei backup
[string[]]$exclude_list = Get-Content -Path "\\itmilitgroup\SD_Utilities\SCRIPT\MigrazioneDati\Automigrazione_exclude_list.log"
$root_path = 'C:\'
$remote_root_list = Get-ChildItem $root_path -Attributes D
$elenco = @();
foreach ($folder in $remote_root_list.Name) {
    if (!($exclude_list -contains $folder)) {
        $elenco += $folder
    }
}  
foreach ($folder in $elenco) {
    if (!($folder -eq "")) {
        if (!($exclude_list -contains $folder)) {
            $full_path = $root_path + $folder
            $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
            $output = [system.String]::Join(" ", $output)
            $output -match "Bytes :\s+(\d+)\s+\d+" > $null
            $size = $Matches[1]
            $backup_list[$folder] = $size
        }
    }
}

# lista delle cartelle da copiare in C:\Users\username
$exclude_list = (
    "AppData",
    "Links",
    "OneDrive",
    "Saved Games",
    "Searches",
    "3D Objects",
    ".cisco",
    ".config"
)
$root_path = "C:\Users\$env:USERNAME\"
$remote_root_list = Get-ChildItem $root_path -Attributes D
foreach ($folder in $remote_root_list.Name) {
    if (!($exclude_list -contains $folder)) {
        $full_path = $root_path + $folder
        $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Bytes :\s+(\d+)\s+\d+" > $null
        $size = $Matches[1]
        $full_path = "Users\$env:USERNAME\$folder"
        $backup_list[$full_path] = $size
    }
}

foreach($key in $backup_list.keys) {  
    Write-Host -NoNewline "C:\$key "
    $size = $backup_list[$key] / 1GB
    $formattato = '{0:0.0}' -f $size
    if ($size -le 5) {
        Write-Host -ForegroundColor Green "<5,0 GB"
    } elseif ($size -le 40) {
        Write-Host -ForegroundColor Yellow "$formattato GB"
    } else {
        Write-Host -ForegroundColor Red "$formattato GB"
    }
}
pause

# Backup wireless network profile
# vedi su https://www.tenforums.com/tutorials/3530-backup-restore-wireless-network-profiles-windows-10-a.html
Write-Host -NoNewline "Backup wireless network profile..."
New-Item -ItemType directory -Path "C:\TEMPSOFTWARE\wifi_profiles" > $null
netsh wlan export profile key=clear folder="C:\TEMPSOFTWARE\wifi_profiles" > $null
Write-Host -ForegroundColor Green " DONE"


<#
# recupero i dati sulla taskbar personalizzata
New-Item -ItemType directory -Path C:\TEMPSOFTWARE > $null
[System.Windows.MessageBox]::Show("Rimuovere dalla taskbar i pin non trasferibili,`nquindi cliccare su Ok",'CONFIGURAZIONE','Ok','Warning') > $null
Copy-Item "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" -Destination "C:\TEMPSOFTWARE" -Force -Recurse > $null
Reg export HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband C:\TEMPSOFTWARE\custom_pinned.reg
#>

# aggiungo l'utente come amministratore locale
Add-LocalGroupMember -Group "Administrators" -Member $env:USERNAME
Write-Host "Local admin attivato" -fore Green

# metto C: in share
$fulluser = $env:UserDomain + "\" + $env:UserName
New-SmbShare –Name "C" –Path "C:\" -FullAccess $fulluser
Write-Host "Unità C: in share" -fore Green

New-Item -ItemType file "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" > $null

"`n*** PC NAME ***`n" | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
$env:computername | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append

"`n*** IP ADDRESS ***`n" | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
(Get-NetIPConfiguration | Where-Object { $_.InterfaceAlias -eq "Ethernet" }).IPv4Address.IPAddress | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append

# lista stampanti e scanner installati
Write-Host "Recupero lista stampanti"

New-Item -ItemType file "C:\TEMPSOFTWARE\usb_devices.log" > $null

"`n*** STAMPANTI ***`n" | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
Get-WMIObject -Class Win32_Printer -Computer  $env:COMPUTERNAME | Select-Object Name,PortName | Where-Object {$_.PortName -match "IP|USB"} | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
Get-WMIObject -Class Win32_Printer -Computer  $env:COMPUTERNAME | Select-Object Name,PortName | Where-Object {$_.PortName -match "IP|USB"} | Out-File "C:\TEMPSOFTWARE\usb_devices.log" -Encoding ASCII -Append
Write-Host "Recupero lista scanner"
"`n*** SCANNER ***`n" | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
$ErrorActionPreference= 'SilentlyContinue'
Get-PnpDevice -Class Image | Select-Object Class,FriendlyName | Where-Object {!($_.FriendlyName -match "Camera|Webcam")} | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
Get-PnpDevice -Class Image | Select-Object Class,FriendlyName | Where-Object {!($_.FriendlyName -match "Camera|Webcam")} | Out-File "C:\TEMPSOFTWARE\usb_devices.log" -Encoding ASCII -Append
$ErrorActionPreference= 'Inquire'

# lista archivi di outlook
Write-Host "Recupero lista archivi Outlook"
if ($OfficeVer -eq '15') {
    $ErrorActionPreference= 'Stop'
    Try {
        $outlook = New-Object -comObject Outlook.Application
        $PST = $outlook.Session.Stores | Where-Object { ($_.FilePath -like '*.PST') }
        New-Item -ItemType file "C:\TEMPSOFTWARE\PST_files.log" > $null
        $PST.FilePath | Out-File "C:\TEMPSOFTWARE\PST_files.log" -Encoding ASCII -Append
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        [System.Windows.MessageBox]::Show("Impossibile recuperare file PST da Outlook",'ERRORE','Ok','Error')
    }
    $ErrorActionPreference= 'Inquire'

    # Blocco Outlook
    Rename-Item -Path "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.EXE" -NewName "C:\Program Files (x86)\Microsoft Office\Office15\OUTLOOK.locked"
    Copy-Item "\\itmilitgroup\SD_Utilities\SCRIPT\MigrazioneDati\OUTLOOK.EXE" -Destination "C:\Program Files (x86)\Microsoft Office\Office15" > $null
} elseif ($OfficeVer -eq '16') {
    # con Office 365 bisogna lanciare l'istanza di creazione dell'oggetto due volte, ignorando l'errore iniziale
    $ErrorActionPreference = 'SilentlyContinue'
    $outlook = New-Object -comObject Outlook.Application

    $ErrorActionPreference= 'Stop'
    Try {
        $outlook = New-Object -comObject Outlook.Application
        $PST = $outlook.Session.Stores | Where-Object { ($_.FilePath -like '*.PST') }
        New-Item -ItemType file "C:\TEMPSOFTWARE\PST_files.log" > $null
        $PST.FilePath | Out-File "C:\TEMPSOFTWARE\PST_files.log" -Encoding ASCII -Append
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        [System.Windows.MessageBox]::Show("Impossibile recuperare file PST da Outlook",'ERRORE','Ok','Error')
    }
    Try {
        $outproc = Get-Process outlook
        Stop-Process -ID $outproc.Id -Force
        Start-Sleep 2
    }
    Catch { 
        [System.Windows.MessageBox]::Show("Assicurarsi che tutte le istanze di Outlook siano chiuse prima di procedere",'TASK MANAGER','Ok','Warning') > $null
    }
    $ErrorActionPreference= 'Inquire'

    # Blocco Outlook
    Rename-Item -Path "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE" -NewName "C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.locked"
    Copy-Item "\\itmilitgroup\SD_Utilities\SCRIPT\MigrazioneDati\OUTLOOK.EXE" -Destination "C:\Program Files\Microsoft Office\root\Office16" > $null
}

# lista dei percorsi di rete
Write-Host "Recupero lista percorsi di rete"
$drives = Get-PSDrive
foreach ($mounted in $drives) {
    if ($mounted.DisplayRoot -match "^\\") {
        New-Item -ItemType file "C:\TEMPSOFTWARE\NetworkDrives.log" > $null
        $string = $mounted.Name + ":;" + $mounted.DisplayRoot
        $string | Out-File "C:\TEMPSOFTWARE\NetworkDrives.log" -Encoding ASCII -Append
    }
}

# lista dei software aggiuntivi installati
Write-Host "Recupero lista software addizionali"

New-Item -ItemType file "C:\TEMPSOFTWARE\optional_sw.log" > $null

"`n*** SOFTWARE ADDIZIONALI ***`n" | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
[string[]]$search_list = Get-Content -Path "\\itmilitgroup\SD_Utilities\SCRIPT\MigrazioneDati\LocalAdmin-Cshared_installed_sw.log"
foreach ($candidate in $search_list) {
    $record = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
          Get-ItemProperty | Where-Object { $_.DisplayName -match $candidate } | Select-Object -Property DisplayName
    $record.DisplayName | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
    $record.DisplayName | Out-File "C:\TEMPSOFTWARE\optional_sw.log" -Encoding ASCII -Append
}
if (Test-Path "C:\Program Files (x86)\Apple Software Update") {
    "SAP BusinessObjects" | Out-File "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log" -Encoding ASCII -Append
    "SAP BusinessObjects" | Out-File "C:\TEMPSOFTWARE\optional_sw.log" -Encoding ASCII -Append
}

# Visualizzo i backup dei MAF di eAudit in locale
$searchpath = "C:\Documenti\eAudIT\Backup"
$today = Get-Date -format "dd/MM/yyyy hh:mm:ss"
if (Test-Path $searchpath) {
    Write-Host -ForegroundColor Cyan "*** BACKUP LOCALI DI EAUDIT ***"
    $folder = Get-ChildItem -Path $searchpath
    foreach ($elem in $folder) {
        if ($elem -match "\.eng$") {
            $backupdate = Get-Item "$searchpath\$elem" | Select-Object -Property fullName, LastWriteTime
            $span = New-TimeSpan -Start $backupdate.LastWriteTime -End $today
            if ($span.Days -gt 30) {
                Write-Host -ForegroundColor Red "$elem -" $backupdate.LastWriteTime
            } elseif ($span.Days -gt 7) {
                Write-Host -ForegroundColor Yellow "$elem -" $backupdate.LastWriteTime
            } else {
                Write-Host -ForegroundColor Green "$elem -" $backupdate.LastWriteTime
            }
        }
    }
    pause
}


# visualizzo il file di log
notepad "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log"
