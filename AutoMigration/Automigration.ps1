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
















# header
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

# Controllo accesso
Import-Module -Name '\\itmilitgroup\SD_Utilities\SCRIPT\Moduli_PowerShell\Patrol.psm1'
$login = Patrol -scriptname Automigrazione

# Importo le funzioni per i form
Import-Module -Name '\\itmilitgroup\SD_Utilities\SCRIPT\Moduli_PowerShell\Forms.psm1'

# directory temporanea
$tmppath = 'C:\TEMPSOFTWARE'
if (!(Test-Path $tmppath)) {
    New-Item -ItemType directory -Path $tmppath > $null
    New-Item -ItemType file "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" > $null
}

# reminder di mettere il PC in GDv7-Clients
[System.Windows.MessageBox]::Show("Spostare $env:COMPUTERNAME su GDv7-Clients",'ATTENZIONE','Ok','Warning')  > $null

# form di scelta modalita'
$form_modalita = FormBase -w 400 -h 230 -text "MODALITA'"
$RadioRETE = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "Esegui migrazione dei dati via rete"
$RadioUSB  = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "Esegui migrazione dei dati via USB"
$RadioNULL = RadioButton -form $form_modalita -checked $false -x 30 -y 80 -text "Non eseguire migrazione dei dati"
OKButton -form $form_modalita -x 90 -y 130 -text "Next"
$result = $form_modalita.ShowDialog()
    
if ($result -eq "OK"){
    if ($RadioRETE.Checked) {
        $modalita = "rete"
        
        $form_IP = FormBase -w 400 -h 200 -text "MODALITA' RETE"
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(350,30)
        $label.Text = "Inserire l'indirizzo IP della vecchia macchina:"
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
    } elseif ($RadioUSB.Checked) {
        $modalita = "usb"

        $form_USB = FormBase -w 400 -h 200 -text "MODALITA' USB"
        $DropDown = new-object System.Windows.Forms.ComboBox
        $DropDown.Location = new-object System.Drawing.Size(10,60)
        $DropDown.Size = new-object System.Drawing.Size(250,30)
        $DropDown.Items.Add('D:')  > $null
        $DropDown.Items.Add('E:')  > $null
        $DropDown.Items.Add('F:')  > $null
        $form_USB.Controls.Add($DropDown)
        $DropDownLabel = new-object System.Windows.Forms.Label
        $DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
        $DropDownLabel.size = new-object System.Drawing.Size(500,30) 
        $DropDownLabel.Text = "Indicare l'unita' associata al disco esterno"
        $form_USB.Controls.Add($DropDownLabel)
        OKButton -form $form_USB -x 100 -y 100 -text "Ok"
        $form_USB.Add_Shown({$DropDown.Select()})
        $result = $form_USB.ShowDialog()
    
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $prefix = $DropDown.Text
        }
    } elseif ($RadioNULL.Checked = $true) {
        $modalita = "null"
        $prefix = "null"
    }
}

# log su \\itmilfsrsd
$itmilfsrsd_logged = $false
if (!($modalita -eq 'null')) {
    net stop workstation /y > $null
    net start workstation > $null
    Start-Sleep 5
    $ErrorActionPreference = 'Stop'
    Try {
        New-PSDrive -Name R -PSProvider FileSystem -Root "\\itmilfsrsd\backupsd_mi\LOG_MIGRAZIONI" -Credential $login > $null
        $itmilfsrsd_logged = $true
    }
    Catch { 
        $errormsg = [System.Windows.MessageBox]::Show("Impossibile accedere alla cartella remota di log",'ATTENZIONE','Ok','Error')       
    }
    $ErrorActionPreference = 'Inquire'
}

Clear-Host
Write-Host -NoNewline "Modalita': "
Write-Host -ForegroundColor Cyan $modalita
Write-Host -NoNewline "Prefix: "
Write-Host -ForegroundColor Cyan $prefix

# verifico la versione di Office
Write-Host -NoNewline "`n`nCheck versione di Office installata... "
$record = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
          Get-ItemProperty | Where-Object {$_.DisplayName -match 'Office' } | Select-Object -Property DisplayName, UninstallString

$OfficeVer = '15'

foreach ($elem in $record) {
    if ($elem -match 'Office 365') {
        $OfficeVer = '16'
    }
}

if ($OfficeVer -eq '16') {
    Write-Host -ForegroundColor Cyan "Installato Office 365"
} else {
    Write-Host -ForegroundColor Cyan "Installato Office 2013"
}

# pannello di controllo
$form_panel = FormBase -w 450 -h 660 -text "STEPS"
$config_sw = CheckBox -form $form_panel -checked $true -x 20 -y 20 -text "Configurazione software presenti"
$install_Chrome = CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "Installazione Chrome"
$install_Chrome.Checked = $false
$serialize_Acrobat = CheckBox -form $form_panel -checked $true -x 20 -y 80 -text "Serializzazione Acrobat"
$install_toolbar = CheckBox -form $form_panel -checked $true -x 20 -y 110 -text "Installazione toolbars"
if ($OfficeVer -eq '16') {
    # le toolbar in Ofice 365 dovrebbero venire su da sole, disabilito l'opzione
    $install_toolbar.Checked = $false
    $install_toolbar.Enabled = $false
}
$upgrade_iTunes = CheckBox -form $form_panel -checked $true -x 20 -y 140 -text "Installazione iTunes aggiornato"
$remove_ConnectedBackup = CheckBox -form $form_panel -checked $true -x 20 -y 170 -text "Rimozione Connected Backup"
$migrazione_dati = CheckBox -form $form_panel -checked $true -x 20 -y 200 -text "Migrazione dati"
if ($modalita -eq 'null') {
    $migrazione_dati.Checked = $false
    $migrazione_dati.Enabled = $false
}
$PST_transfer = CheckBox -form $form_panel -checked $true -x 20 -y 230 -text "Trasferimento archivi di Outlook"
if (!($modalita -eq 'rete')) {
    $PST_transfer.Checked = $false
    $PST_transfer.Enabled = $false
}
$network_map = CheckBox -form $form_panel -checked $true -x 20 -y 260 -text "Mappatura percorsi di rete"
if (!($modalita -eq 'rete')) {
    $network_map.Checked = $false
    $network_map.Enabled = $false
}
$setting_taskbar = CheckBox -form $form_panel -checked $true -x 20 -y 290 -text "Settaggio taskbar"
$BitLocker_services = CheckBox -form $form_panel -checked $true -x 20 -y 320 -text "Attivazione servizi BitLocker"
$machine_certificate = CheckBox -form $form_panel -checked $true -x 20 -y 350 -text "Configurazione certificato macchina"
$wifi_profiles = CheckBox -form $form_panel -checked $true -x 20 -y 380 -text "Importazione profili WiFi"
if (!($modalita -eq 'rete')) {
    $wifi_profiles.Checked = $false
    $wifi_profiles.Enabled = $false
}
$optional_sw = CheckBox -form $form_panel -checked $true -x 20 -y 410 -text "Installazione software opzionali"
if (!($modalita -eq 'rete')) {
    $optional_sw.Checked = $false
    $optional_sw.Enabled = $false
}
$optional_usb = CheckBox -form $form_panel -checked $true -x 20 -y 440 -text "Installazione periferiche USB"
if (!($modalita -eq 'rete')) {
    $optional_usb.Checked = $false
    $optional_usb.Enabled = $false
}
$sendlog = CheckBox -form $form_panel -checked $true -x 20 -y 470 -text "Invio mail di LOG"
# tengo l'opzione disabilitata in attesa di news
$sendlog.Checked = $false
$sendlog.Enabled = $false


OKButton -form $form_panel -x 100 -y 550 -text "Ok"
$result = $form_panel.ShowDialog()

# checks
if (!($prefix -eq "null")) {
    if (!(Test-Path $prefix)) {
        [System.Windows.MessageBox]::Show("$prefix non connesso",'ERRORE','Ok','Error') > $null
        Exit
    }
}

$check = Test-Connection "itmilitgroup" -quiet
if (!($check)) {
    [System.Windows.MessageBox]::Show("itmilitgroup non connesso",'ERRORE','Ok','Error') > $null
    Exit
}


if ($modalita -eq 'rete') {
    $log_pcvecchio = $prefix + "\TEMPSOFTWARE\LocalAdmin-Cshared.log"
    Copy-Item $log_pcvecchio -Destination "C:\TEMPSOFTWARE" > $null
    $logcontent = Get-Content "C:\TEMPSOFTWARE\LocalAdmin-Cshared.log"
    $pcname = $logcontent[3]
    "*** MIGRAZIONE DATI DA: $pcname ***" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
} elseif ($modalita -eq 'usb') {
    $form_EXP = FormBase -w 300 -h 150 -text "NOME PC"
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(150,30)
    $label.Text = "Nome PC vecchio:"
    $form_EXP.Controls.Add($label)
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(160,20)
    $textBox.Size = New-Object System.Drawing.Size(120,30)
    $form_EXP.Controls.Add($textBox)
    OKButton -form $form_EXP -x 100 -y 50 -text "Ok"
    $form_EXP.Add_Shown({$textBox.Select()})
    $result = $form_EXP.ShowDialog()
    $pcname = $textBox.Text
    "*** MIGRAZIONE DATI DA: $pcname ***" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

#lancio i singoli moduli selezionati dal pannello di controllo
if ($config_sw.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Configurazione software presenti ***"

    # Autoconfigurazione silent di Outlook
    [System.Windows.MessageBox]::Show("Configurazione Outlook...",'CONFIGURAZIONE SW','Ok','Info') > $null
    Write-Host -NoNewline "Configurazione Outlook..."
    
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

    # Autoconfigurazione silent del LanguagePack di Office
    # per la configurazione manuale lanciare il file C:\Program Files (x86)\Microsoft Office\Office15\SETLANG.EXE
    [System.Windows.MessageBox]::Show("Settaggio lingua Office...",'CONFIGURAZIONE SW','Ok','Info') > $null
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
    
    Start-Sleep 20
    Write-Host -ForegroundColor Green " DONE"

    $lista_startup = [ordered]@{
        "Configurazione Skype..."              = "lync"
        "Configurazione WinZip..."             = "winzip"
    }

    foreach($key in $lista_startup.keys) {  
        [System.Windows.MessageBox]::Show($key,'CONFIGURAZIONE SW','Ok','Info') > $null
        Write-Host -NoNewline "$key"
        $startup_path = $lista_startup[$key]
        Start-Process $startup_path
        Start-Sleep 20
        Write-Host -ForegroundColor Green " DONE"
    }

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Configurazione software presenti" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($install_Chrome.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Installazione Chrome ***"
    [System.Windows.MessageBox]::Show("Installazione Chrome...",'INSTALLAZIONE SW','Ok','Info') > $null

    $source_path = "\\itmilitgroup\sd_utilities\UtilitiesSD\Preparazione_PC_Kpmg\MigrazionePC\Chrome"
    $dest_path = "C:\TEMPSOFTWARE"
    $opts = ("/E", "/Z", "/NP")
    $cmd_args = ($source_path, $dest_path, $opts)
    Write-Host -NoNewline "Recupero launcher di Chrome..."
    robocopy @cmd_args > $null
    Write-Host -ForegroundColor Green " DONE"

    Write-Host -NoNewline "Lancio l'installer di Chrome..."
    Start-Process "C:\TEMPSOFTWARE\ChromeSetup.exe"
    Start-Sleep 20
    Write-Host -ForegroundColor Green " DONE"

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Installazione Chrome" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($serialize_Acrobat.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Serializzazione Acrobat ***"

    Write-Host -NoNewline "Recupero Update_Adobe_Serialize..."
    $dest_path = "C:\TEMPSOFTWARE\Update_Adobe_Serialize"
    $opts = ("/E", "/Z", "/NP")
    $cmd_args = ("\\itmilitgroup\SD_Utilities\SCRIPT\Adobe_Acrobat\Adobe_Serialize", $dest_path, $opts)
    robocopy @cmd_args > $null
    Write-Host -ForegroundColor Green " DONE"

    Write-Host -NoNewline "Serializzazione Acrobat..."
    Start-Process "C:\TEMPSOFTWARE\Update_Adobe_Serialize\Serialize.vbs"
    Write-Host -ForegroundColor Green " DONE"

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Serializzazione Acrobat" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($install_toolbar.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Installazione toolbars ***"

    $swlist = [ordered]@{
        excel = @{
            source_path = "\\itmilitgroup\sd_utilities\UtilitiesSD\Preparazione_PC_Kpmg\MigrazionePC\Office Toolbar 2013\Toolbar EXCEL"
            desc = "Excel toolbar..."
            installer = "C:\TEMPSOFTWARE\KPMG_Magic_Excel_Plug-in_5.1.msi"
        }
        
        powerpoint = @{
            source_path = "\\itmilitgroup\sd_utilities\UtilitiesSD\Preparazione_PC_Kpmg\MigrazionePC\Office Toolbar 2013\Toolbar PPT"
            desc = "Powerpoint toolbar..."
            installer = "C:\TEMPSOFTWARE\KPMG_Global_PowerPoint_Toolbar_A4_5.1.msi"
        }
        
        outlook = @{
            source_path = "\\itmilitgroup\sd_utilities\UtilitiesSD\Preparazione_PC_Kpmg\MigrazionePC\PROOFPOINT TOOLBAR"
            desc = "Proofpoint toolbar..."
            installer = "C:\TEMPSOFTWARE\Setup.vbs"
        }
    }

    foreach ($key in $swlist.Keys) {

        [System.Windows.MessageBox]::Show("Installazione " + $swlist.$key.desc,'INSTALLAZIONE SW','Ok','Info') > $null

        Write-Host -NoNewline "Recupero" $swlist.$key.desc
        $dest_path = "C:\TEMPSOFTWARE"
        $opts = ("/E", "/Z", "/NP")
        $cmd_args = ($swlist.$key.source_path, $dest_path, $opts)
        robocopy @cmd_args > $null
        Write-Host -ForegroundColor Green " DONE"
        
        # chiudo Outlook prima dell'installazione del plugin ProofPoint
        if ($key -eq 'outlook') {
            $ErrorActionPreference = 'Stop'
            Try {
                Get-Process outlook | Foreach-Object { $_.CloseMainWindow() | Out-Null }
                Start-Sleep 2
            }
            Catch { 
                [System.Windows.MessageBox]::Show("Non trovo il processo di Outlook,`nprobabilmente Outlook e' gia' chiuso",'TASK MANAGER','Ok','Info') > $null
            }
            $ErrorActionPreference = 'Inquire'
        }

        Write-Host -NoNewline "Installo" $swlist.$key.desc
        if ($key -eq 'outlook') {
            Start-Process $swlist.$key.installer
        } else {
            msiexec /i $swlist.$key.installer /passive
        }
        Start-Sleep 20
        Write-Host -ForegroundColor Green " DONE"
    }

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Installazione toolbars" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($upgrade_iTunes.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Installazione iTunes aggiornato ***"

    [System.Windows.MessageBox]::Show("Upgrade iTunes...",'AGGIORNAMENTO SW','Ok','Info') > $null
    Write-Host -NoNewline "Rimozione iTunes..."
    $record = Get-ChildItem -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall, HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall |
              Get-ItemProperty | Where-Object {$_.DisplayName -match "iTunes" } | Select-Object -Property DisplayName, UninstallString
    $record -match "MsiExec\.exe /I\{([A-Z0-9\-]+)\}" > $null
    $string = $Matches[1] 
    msiexec "/X{$string}" /passive
    [System.Windows.MessageBox]::Show("Clicca OK quando la disinstallazione sara' terminata",'ATTENDI','Ok','Info') > $null
    Write-Host -ForegroundColor Green " DONE"
    Write-Host -NoNewline "Download iTunes aggiornato..."
    $download_itunes = New-Object net.webclient
    $download_itunes.Downloadfile("https://www.apple.com/itunes/download/win64", "C:\TEMPSOFTWARE\updated_iTunes.exe")
    Write-Host -ForegroundColor Green " DONE"
    Write-Host -NoNewline "Installazione iTunes aggiornato..."
    Start-Process -Wait "C:\TEMPSOFTWARE\updated_iTunes.exe" /q
    Start-Sleep 5
    Remove-Item "C:\TEMPSOFTWARE\updated_iTunes.exe" -Force
    Write-Host -ForegroundColor Green " DONE"

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Upgrade iTunes" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($remove_ConnectedBackup.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Rimozione Connected Backup ***"

    [System.Windows.MessageBox]::Show("Rimozione Connected Backup...",'CONFIGURAZIONE SW','Ok','Info') > $null
    
    Write-Host -NoNewline "Rimozione Connected Backup..."
    msiexec "/X{393E4C89-67E9-43BF-AD29-94D19F7624F7}" /passive
    Start-Sleep 30
    Write-Host -ForegroundColor Green " DONE"

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Rimozione Connected Backup" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($migrazione_dati.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Migrazione dati ***"

    $ErrorActionPreference= 'SilentlyContinue'

    # copio i bookmarks di Chrome
    if ($install_Chrome.Checked -eq $false) {
        Start-Process -Wait "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    }
    $bookmarks = $prefix + "\Users\$env:USERNAME\AppData\Local\Google\Chrome\User Data\Default\Bookmarks"
    if (Test-Path $bookmarks -PathType Leaf) {
        Write-Host -NoNewline "Copio bookmarks di Chrome..."
        Copy-Item $bookmarks -Destination "C:\Users\$env:USERNAME\AppData\Local\Google\Chrome\User Data\Default" > $null
        Write-Host -ForegroundColor Green " DONE"
    }

    # Copio i preferiti di Outlook
    $outlook_aspect = $prefix + "\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook\Outlook.xml"
    if (Test-Path $outlook_aspect -PathType Leaf) {
        Write-Host -NoNewline "Copio aspetto di Outlook..."
        Remove-Item "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook\Outlook.xml" -Force
        Copy-Item $outlook_aspect -Destination "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Outlook" > $null
        Write-Host -ForegroundColor Green " DONE"
    }

    # copio le icone KPMG
    $iconpath = $prefix + "\Program Files\KPMG\Global Desktop\Utilities"
    if (Test-Path $iconpath) {
        Write-Host -NoNewline "Copio icone di KPMG..."
        robocopy $iconpath "C:\Program Files\KPMG\Global Desktop\Utilities" /E /Z /NP > $null
        Write-Host -ForegroundColor Green " DONE"
    }

    Write-Host -NoNewline "Cerco i percorsi di backup..."
    $backup_list = @{} # variabili in cui inseriro' i percorsi su cui far la migrazione

    # lista cartelle su cui si dovrebbe fare la migrazione; se la cartella non e' vuota la metto nella lista dei backup    
    [string[]]$allow_list = Get-Content -Path "\\itmilitgroup\SD_Utilities\SCRIPT\MigrazioneDati\Automigrazione_allow_list.log"
    $allow_list = $allow_list -replace ('\$username', $env:USERNAME)             
    foreach ($folder in $allow_list) {
        $full_path = $prefix + '\' + $folder
        if (Test-Path $full_path) {
            $output = robocopy $full_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
            $output = [system.String]::Join(" ", $output)
            $output -match "Bytes :\s+(\d+)\s+\d+" > $null
            $size = $Matches[1]
            if ($size -gt 1KB) {
                $backup_list[$folder] = $size
                Write-Host -NoNewline "."
            }
        }
    }

    # lista cartelle da escludere dalla migrazione; se in C:\ ci sono altre cartelle oltre a queste le includo nella lista dei backup
    [string[]]$exclude_list = Get-Content -Path "\\itmilitgroup\SD_Utilities\SCRIPT\MigrazioneDati\Automigrazione_exclude_list.log"
    $root_path = $prefix + '\'
    $remote_root_list = Get-ChildItem $root_path -Attributes D
    $elenco = @();
    foreach ($folder in $remote_root_list.Name) {
        if (!($exclude_list -contains $folder)) {
            $elenco += $folder
        }
    }
    $string = [system.String]::Join("`r`n", $elenco)
    $form_folders = FormBase -w 400 -h 275 -text "ELENCO CARTELLE IN C:\"
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(350,30)
    $label.Text = "Cancella le cartelle che non vuoi trasferire:"
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
                $output -match "Bytes :\s+(\d+)\s+\d+" > $null
                $size = $Matches[1]
                $backup_list[$folder] = $size
                Write-Host -NoNewline "."
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
    $root_path = "$prefix\Users\$env:USERNAME\"
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
            Write-Host -NoNewline "."
        }
    }

    Write-Host -ForegroundColor Green " DONE"

    # blocco del singolo job di migrazione
    $RoboCopyBlock = {
        param($final_path,$prefisso)
        $filename = $final_path.Replace('\','-')
        if (Test-Path "C:\TEMPSOFTWARE\ROBOCOPY_$filename.log" -PathType Leaf) {
            Remove-Item "C:\TEMPSOFTWARE\ROBOCOPY_$filename.log" -Force
        }
        New-Item -ItemType file "C:\TEMPSOFTWARE\ROBOCOPY_$filename.log" > $null
        $source_path = $prefisso + '\' + $final_path
        $dest_path = 'C:\' + $final_path
        $opts = ("/E", "/Z", "/NP", "/W:5", "/R:5", "/V", "/LOG+:C:\TEMPSOFTWARE\ROBOCOPY_$filename.log")
        $cmd_args = ($source_path, $dest_path, $opts)    
        robocopy @cmd_args
    }

    # lancio i job di migrazione in parallelo
    $Time = [System.Diagnostics.Stopwatch]::StartNew()
    foreach ($folder in $backup_list.Keys) {
        Write-Host -NoNewline -ForegroundColor Cyan "$folder"
        Start-Job $RoboCopyBlock -Name $folder -ArgumentList $folder, $prefix > $null
        Write-Host -ForegroundColor Green " JOB STARTED"
    }

    Start-Sleep 10

    # blocco per disegnare la progress bar
    $form_bar = New-Object System.Windows.Forms.Form
    $form_bar.Text = "TRASFERIMENTO DATI"
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

    # Attendo che i job vengano completati mostrando un progress
    While (Get-Job -State "Running") {
        Clear-Host
        Write-Host -ForegroundColor Yellow "*** TRASFERIMENTO DATI IN CORSO ***"

        $total_bytes = 0
        $trasferred_bytes = 0
        
        foreach ($folder in $backup_list.Keys) {
            Write-Host -NoNewline "C:\$folder "
            $source_path = $prefix + '\' + $folder
            $source_size = $backup_list[$folder]
            $dest_path = 'C:\' + $folder
            $output = robocopy $dest_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
            $output = [system.String]::Join(" ", $output)
            $output -match "Bytes :\s+(\d+)\s+\d+" > $null
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
        $label.Text = "Avanzamento totale: $formattato% - $estimated minuti alla fine"
        $bar.Value = $progress
        $form_bar.Refresh()

        Write-Host " "
        Start-Sleep 5
    }

    $form_bar.Close()

    $joblog = Get-Job | Receive-Job # Recupero l'output dai job
    Remove-Job * # Cleanup

    # controllo dimensioni
    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Migrazione dati" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
    Write-Host " "
    Write-Host -NoNewline "Controllo dimensioni..."
    foreach ($folder in $backup_list.Keys) {
        $source_path = $prefix + '\' + $folder
        $source_size = $backup_list[$folder]
        $dest_path = 'C:\' + $folder
        $output = robocopy $dest_path c:\fakepath /L /XJ /R:0 /W:1 /NP /E /BYTES /NFL /NDL /NJH /MT:64
        $output = [system.String]::Join(" ", $output)
        $output -match "Bytes :\s+(\d+)\s+\d+" > $null
        $dest_size = $Matches[1]

        $foldername = $folder.Replace('\','-')                      
        if ($dest_size -ge $source_size) { # la copia e' andata a buon fine
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
            # ALERT: in caso di anomalie
            if ($itmilfsrsd_logged -eq $true) {
                $tocheck = Get-Content "C:\TEMPSOFTWARE\ROBOCOPY_$foldername.log" | Select-Object -Last 12
                foreach ($newline in $tocheck) {
                    if ($newline -match "Files :") {
                        $newline -match "Files :\s+\d+\s+\d+\s+(\d+\s+\d+\s+\d+\s+\d+)$" > $null
                        $string = $Matches[1]
                        if ($string -notmatch "0         0         0         0") {
                            $WarningPreference = 'SilentlyContinue'
                            $ErrorActionPreference = 'Stop'
                            Try {
                                $alertpath = "R:\$env:USERNAME-$env:COMPUTERNAME"
                                if (!(Test-Path $alertpath)) {
                                    New-Item -ItemType directory -Path $alertpath > $null
                                }
                                Copy-Item "C:\TEMPSOFTWARE\ROBOCOPY_$foldername.log" -Destination $alertpath > $null
                            }
                            Catch {
                                Write-Host -ForegroundColor Red "`nImpossibile copiare file di log"
                            }
                            $ErrorActionPreference = 'Inquire'
                            $WarningPreference = 'Inquire'
                        }
                    }
                } 
            } else {
                [System.Windows.MessageBox]::Show("Controllare il file C:\TEMPSOFTWARE\ROBOCOPY_$foldername.log",'ATTENZIONE','Ok','Warning')
            }

        } else { # la copia NON e' andata a buon fine
            # ALERT: in caso di anomalie
            if ($itmilfsrsd_logged -eq $true) {
                $ErrorActionPreference = 'Stop'
                Try {
                    $alertpath = "R:\$env:USERNAME-$env:COMPUTERNAME"
                    if (!(Test-Path $alertpath)) {
                        New-Item -ItemType directory -Path $alertpath > $null
                    }
                    Copy-Item "C:\TEMPSOFTWARE\ROBOCOPY_$foldername.log" -Destination $alertpath > $null
                }
                Catch {
                    Write-Host -ForegroundColor Red "`nImpossibile copiare file di log"
                }
                $ErrorActionPreference = 'Inquire'
            } else {
                [System.Windows.MessageBox]::Show("Controllare il file C:\TEMPSOFTWARE\ROBOCOPY_$foldername.log",'ATTENZIONE','Ok','Warning')
            }
                                   
            Clear-Host
            $diff = $source_size - $dest_size
            Write-Host "PATH.........: $folder`nSOURCE SIZE..: $source_size bytes`nDEST SIZE....: $dest_size bytes`nDIFF SIZE....: $diff bytes"

            $whatif = [System.Windows.MessageBox]::Show("La copia di $folder non e' andata a buon fine.`nVedere il file di log?",'ERRORE','YesNo','Error')
            if ($whatif -eq "Yes") {
                notepad "C:\TEMPSOFTWARE\ROBOCOPY_$foldername.log"
            }
            $whatif = [System.Windows.MessageBox]::Show("La copia di $folder non e' andata a buon fine.`nRilanciare la copia?",'ERRORE','YesNo','Error')
            if ($whatif -eq "Yes") {
                New-Item -ItemType file "C:\TEMPSOFTWARE\ROBOCOPY-RETRY_$foldername.log" > $null
                $opts = ("/E", "/Z", "/NP", "/W:5", "/V", "/TEE", "/LOG+:C:\TEMPSOFTWARE\ROBOCOPY-RETRY_$foldername.log")
                $cmd_args = ($source_path, $dest_path, $opts)    
                Write-Host -ForegroundColor Yellow "RETRY: copia di $folder in corso..."
                Start-Sleep 3
                robocopy @cmd_args
                notepad "C:\TEMPSOFTWARE\ROBOCOPY-RETRY_$foldername.log"
                $whatif = [System.Windows.MessageBox]::Show("La copia di $folder e' andata a buon fine?",'CONFERMA','YesNo','Info')                
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
                    [System.Windows.MessageBox]::Show("Copiare $folder manualmente",'CONFERMA','Ok','Info') > $null
                }
            }
        }    
    }
    $ErrorActionPreference= 'Inquire'
    Write-Host -ForegroundColor Green " DONE"

    # check attributi su cartelle nascoste
    Write-Host " "
    Write-Host -NoNewline "Check attributi..."
    foreach ($folder in $backup_list.Keys) {
        $dest_path = 'C:\' + $folder
        attrib -s -h $dest_path
    }
    Write-Host -ForegroundColor Green " DONE"

    # Copia dei file nella cartella Users
    Write-Host -NoNewline "Copia files in C:\Users\$env:USERNAME..."
    $userfiles = Get-ChildItem "C:\Users\$env:USERNAME" -Attributes A
    foreach ($afile in $userfiles) {
        Copy-Item "$prefix\Users\$env:USERNAME\$afile" -Destination "C:\Users\$env:USERNAME" -Force > $null
    }
    Write-Host -ForegroundColor Green " DONE"

    # Check di OneNote 365
    if (Test-Path -Path "C:\Users\$env:USERNAME\AppData\Local\Packages\Microsoft.Office.OneNote_8wekyb3d8bbwe") {
        [System.Windows.MessageBox]::Show("Importare i notebook per OneNote",'ATTENZIONE','Ok','Info') > $null
    }
    
    # Copia forzata dei file di Sticky Notes
    $source_path = $prefix + "\Users\$env:USERNAME\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe\LocalState"
    $dest_path = "C:\Users\$env:USERNAME\AppData\Local\Packages\Microsoft.MicrosoftStickyNotes_8wekyb3d8bbwe"
    if (Test-Path $source_path) {
        Write-Host -NoNewline "Trasferimento Sticky Notes..."
        [System.Windows.MessageBox]::Show("Lanciare Sticky Notes e poi chiuderlo.`nClicca so Ok dopo aver chiuso Sticky Notes",'ATTENZIONE','Ok','Info') > $null
        Remove-Item "$dest_path\LocalState" -Force -Recurse
        Start-Sleep 3
        Copy-Item $source_path -Destination $dest_path -Force -Recurse > $null
        Write-Host -ForegroundColor Green " DONE"
    }
}

if ($PST_transfer.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Trasferimento archivi di Outlook ***"

    $pst_file_sourcepath = $prefix + "\TEMPSOFTWARE\PST_files.log"
    Write-Host -NoNewline "Recupero la lista dei file PST..."
    Copy-Item $pst_file_sourcepath -Destination "C:\Optional Software" > $null
    Write-Host -ForegroundColor Green " DONE"

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Trasferimento archivi di Outlook" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($network_map.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Mappatura percorsi di rete ***"

    Write-Host "Mappatura percorsi di rete"
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
        Write-Host -ForegroundColor Cyan "Nessun percorso trovato"
    }

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Mappatura percorsi di rete" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($setting_taskbar.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Settaggio taskbar ***"

    Write-Host -NoNewline "Settaggio taskbar..."

    # settaggio taskbar standard
    Remove-Item "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" -Recurse -Force
    Remove-Item "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband" -Recurse -Force
    robocopy "\\itmilitgroup\SD_Utilities\UtilitiesSD\Preparazione_PC_Kpmg\MigrazionePC\PIN" "C:\TEMPSOFTWARE\PIN" /E /Z /NP > $null
    robocopy "\\itmilitgroup\SD_Utilities\UtilitiesSD\Preparazione_PC_Kpmg\MigrazionePC\PIN" "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar" /E /Z /NP > $null
    Remove-Item "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\PIN.reg" -Force
    Remove-Item "C:\Users\$env:USERNAME\AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar\old" -Force -Recurse
    Start-Process "C:\TEMPSOFTWARE\PIN\PIN.reg" -Wait

<#    
    # settaggio taskbar personalizzata
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
#>
    
    Write-Host -ForegroundColor Green " DONE"

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Settaggio Taskbar" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($BitLocker_services.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Attivazione servizi BitLocker ***"

    "BDESVC", "MBAMAgent" | foreach {
        Set-Variable -Name Service_Name -Value $_
        if ((Get-Service -Name $Service_Name).Status -eq "Stopped") {
            Set-Service -Name $Service_Name -StartupType Automatic
            Start-Service -Name $Service_Name
        }
        $out = Get-Service -Name $Service_Name
        Write-Host -ForegroundColor Cyan "$($out.Displayname)`t`t$($out.Status)`t($($out.Starttype))"
    }

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Attivazione servizi BitLocker" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($machine_certificate.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Configurazione certificato macchina ***"

    $cert = Get-ChildItem -Path cert:\LocalMachine\My | Where-Object {$_.Subject -match 'it.kworld.kpmg.com'}
    if ($cert -ne $null -and $cert.PrivateKey.CspKeyContainerInfo.UniqueKeyContainerName -ne $null) {
        $keyPath = $env:ProgramData + "\Microsoft\Crypto\RSA\MachineKeys\"; 
        $keyName = $cert.PrivateKey.CspKeyContainerInfo.UniqueKeyContainerName;
        $keyFullPath = $keyPath + $keyName;

        Write-Host "Granting NETWORK SERVICE..." -NoNewline
        $acl = (Get-Item $keyFullPath).GetAccessControl('Access')
    
        $buildAcl = New-Object  System.Security.AccessControl.FileSystemAccessRule('NETWORK SERVICE','Read','Allow')
        $acl.SetAccessRule($buildAcl)
        Set-Acl $keyFullPath $acl

        $buildAcl = New-Object  System.Security.AccessControl.FileSystemAccessRule('NETWORK SERVICE','FullControl','Allow')
        $acl.SetAccessRule($buildAcl)
        Set-Acl $keyFullPath $acl

        Write-Host " DONE" -ForegroundColor Green
    } else {
        Write-Host -ForegroundColor Red "FAIL"
        Write-Host "Impossibile trovare il certificato"
    }

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Configurazione certificato macchina" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($wifi_profiles.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Importazione profili WiFi ***"

    $wifi_profile_sourcepath = $prefix + "\TEMPSOFTWARE\wifi_profiles"
    Copy-Item -Recurse $wifi_profile_sourcepath -Destination "C:\TEMPSOFTWARE\" > $null
    
    Write-Host "Restore wireless network profile"
    $wifi_profile_list = Get-ChildItem "C:\TEMPSOFTWARE\wifi_profiles"
    $exclude_list = ("Wi-Fi-ITSWireless.xml", "Wi-Fi-KLogin.xml")
    foreach ($wifi_profile in $wifi_profile_list) {
        if ($exclude_list -contains $wifi_profile) {
            Write-Host "Skip $wifi_profile"
        } else {
            netsh wlan add profile filename="C:\TEMPSOFTWARE\wifi_profiles\$wifi_profile" user=current
        }
    }
    Remove-Item "C:\TEMPSOFTWARE\wifi_profiles" -Recurse -Force

    $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
    "$date Importazione profili WiFi" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
}

if ($optional_sw.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Installazione software opzionali ***"

    $optsw_sourcepath = $prefix + "\TEMPSOFTWARE\optional_sw.log"
    Copy-Item $optsw_sourcepath -Destination "C:\TEMPSOFTWARE" > $null

    $ErrorActionPreference = 'SilentlyContinue'
    $opt_sw_list = Get-Content -Path "C:\TEMPSOFTWARE\optional_sw.log"
    $opt_sw_string = [system.String]::Join(";", $opt_sw_list)
    $ErrorActionPreference= 'Inquire'

    if ($OfficeVer -eq '16') { # Visio e Project 2019
        $xml_conf = @{
            'Project Standard' = '"C:\TEMPSOFTWARE\Install_Office365_v1908\ProjectStdXVolume.xml"'
            'Project Professional' = '"C:\TEMPSOFTWARE\Install_Office365_v1908\ProjectProXVolume.xml"'
            'Visio Standard' = '"C:\TEMPSOFTWARE\Install_Office365_v1908\VisioStdXVolume.xml"'
            'Visio Professional' = '"C:\TEMPSOFTWARE\Install_Office365_v1908\VisioProXVolume.xml"'
        }
        foreach ($key in $xml_conf.Keys) {
            if ($opt_sw_string -match $key) {
                [System.Windows.MessageBox]::Show("Installazione $key",'CONFIGURAZIONE SW','Ok','Info') > $null

                $source_path = '"\\itmilitgroup\SD_Utilities\GDStandard\Install_Office365_v1908"'
                $dest_path = '"C:\TEMPSOFTWARE\Install_Office365_v1908"'
                Write-Host -NoNewline "Recupero Office 365"
                $string = "$source_path $dest_path /E /Z"
                Start-Process -Wait robocopy $string
                Write-Host -ForegroundColor Green " DONE"

                $opt_string = ('/Configure', $xml_conf[$key])
                Write-Host -NoNewline "Installo $key"
                Start-Process -Wait "C:\TEMPSOFTWARE\Install_Office365_v1908\setup.exe" $opt_string
                Write-Host -ForegroundColor Green " DONE"

                $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
                "$date Installazione $key" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append  
            }
        }
    } else { # Visio e Project 2013
        $opt_list = @{
            Project = @{
                string = 'Project Standard'
                source_path = '"\\itmilitgroup\SD_Utilities\UtilitiesSD\Software\Microsoft Project 2013"'
                dest_path = '"C:\TEMPSOFTWARE\Microsoft Project 2013"'
                exe_path = 'C:\TEMPSOFTWARE\Microsoft Project 2013\Setup.exe'
            }
            Visio = @{
                string = 'Visio Standard'
                source_path = '"\\itmilitgroup\SD_Utilities\UtilitiesSD\Software\Microsoft Visio 2013"'
                dest_path = '"C:\TEMPSOFTWARE\Microsoft Visio 2013"'
                exe_path = 'C:\TEMPSOFTWARE\Microsoft Visio 2013\Setup.exe'
            }
        }

        foreach ($key in $opt_list.Keys) {
            if ($opt_sw_string -match $opt_list[$key].string) {
                [System.Windows.MessageBox]::Show("Installazione $key",'CONFIGURAZIONE SW','Ok','Info') > $null

                Write-Host -NoNewline "Recupero" $key
                $source = $opt_list[$key].source_path
                $dest = $opt_list[$key].dest_path
                $string = "$source $dest /E /Z"
                Start-Process -Wait robocopy $string
                Write-Host -ForegroundColor Green " DONE"
        
                Write-Host -NoNewline "Installo" $key
                Start-Process $opt_list[$key].exe_path
                Start-Sleep 20
                Write-Host -ForegroundColor Green " DONE"

                $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
                "$date Installazione $key" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
            }
        }
    }

    if ($opt_sw_string -match 'SAP GUI') {
        [System.Windows.MessageBox]::Show("Installazione SAP GUI",'CONFIGURAZIONE SW','Ok','Info') > $null

        Write-Host -NoNewline "Recupero SAP GUI"
        $source = '"\\itmilitgroup\SD\Software\SapGUI_760"'
        $dest = '"C:\Optional Software\SapGUI_760"'
        $string = "$source $dest /E /Z"
        Start-Process -Wait robocopy $string
        Write-Host -ForegroundColor Green " DONE"
        
        Write-Host -NoNewline "Installo SAP GUI"
        Start-Process -Wait 'C:\Optional Software\SapGUI_760\BD_NW_7.0_Presentation_7.60_Comp._1_\PRES1\GUI\WINDOWS\Win32\SetupAll.exe'
        Write-Host -ForegroundColor Green " DONE"

        $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
        "$date Installazione SAP GUI" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
    }

    if ($opt_sw_string -match 'IDEA') {
        [System.Windows.MessageBox]::Show("Installazione IDEA 10.3",'CONFIGURAZIONE SW','Ok','Info') > $null
        if ($opt_sw_string -match 'KPMG eA201') {
            Write-Host -ForegroundColor Cyan "IDEA gia' installato"
        } else {
            Write-Host -NoNewline "Installo IDEA 10.3"
            Set-Location -Path "C:\Optional Software\IDEA 10.3"
            Start-Process "C:\Optional Software\IDEA 10.3\Setup.exe"
            Start-Sleep 20
            Write-Host -ForegroundColor Green " DONE"
        }

        $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
        "$date Installazione IDEA" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
    }
    
    if ($opt_sw_string -match 'Dike') {
        [System.Windows.MessageBox]::Show("Installazione Dike 6",'CONFIGURAZIONE SW','Ok','Info') > $null

        Write-Host -NoNewline "Recupero Dike 6"
        Copy-Item "\\itmilitgroup\SD_Utilities\UtilitiesSD\Osama\Utility\DIke6-installer-win32.msi" -Destination "C:\TEMPSOFTWARE\DIke6-installer-win32.msi" > $null
        Write-Host -ForegroundColor Green " DONE"

        Write-Host -NoNewline "Installo Dike 6"
        Start-Process "C:\TEMPSOFTWARE\DIke6-installer-win32.msi"
        Start-Sleep 20
        [System.Windows.MessageBox]::Show("Andare alla sezione 'Impostazioni' (logo ingranaggio) ed impostare proxy itmilprx18.it.kworld.kpmg.com e porta 8080. Inserire nome utente e password di Windows.",'CONFIGURAZIONE Dike 6','Ok','Info') > $null
        Write-Host -ForegroundColor Green " DONE"

        $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
        "$date Installazione Dike" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
    }
    
    if ($opt_sw_string -match 'Power BI') {
        [System.Windows.MessageBox]::Show("Installazione Power BI",'CONFIGURAZIONE SW','Ok','Info') > $null

        Write-Host -NoNewline "Recupero Power BI"
        Copy-Item "\\itmilitgroup\SD_Utilities\UtilitiesSD\Osama\Power BI Desktop win 10 64 BIT.msi" -Destination "C:\TEMPSOFTWARE\Power BI Desktop win 10 64 BIT.msi" > $null
        Write-Host -ForegroundColor Green " DONE"

        Write-Host -NoNewline "Installo Power BI"
        Start-Process "C:\TEMPSOFTWARE\Power BI Desktop win 10 64 BIT.msi"
        Start-Sleep 20
        Write-Host -ForegroundColor Green " DONE"

        $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
        "$date Installazione Power BI" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
    }
    
    $upgrade_sap = 'no'
    if ($opt_sw_string -match 'SAP BusinessObjects') {
        if ($OfficeVer -eq '16') {
            # se c'e' Office 365 metto il plugin aggiornato
            $upgrade_sap = 'yes'
        } else {
            [System.Windows.MessageBox]::Show("Installazione SAP BusinessObjects",'CONFIGURAZIONE SW','Ok','Info') > $null

            Write-Host -NoNewline "Recupero SAP BusinessObjects"
            $source = '"\\itmilitgroup\SD_Utilities\UtilitiesSD\SAP BusinessObjects"'
            $dest = '"C:\Program Files (x86)\SAP BusinessObjects"'
            $string = "$source $dest /E /Z"
            Start-Process -Wait robocopy $string
            Copy-Item -Recurse "\\itmilitgroup\sd_utilities\UtilitiesSD\Utility\SapBusinessChiavediRegistro\SAPBUSINESS.reg" -Destination  "C:\TEMPSOFTWARE" > $null
            Write-Host -ForegroundColor Green " DONE"
        
            Write-Host -NoNewline "Installo SAP BusinessObjects"
            Start-Process "C:\TEMPSOFTWARE\SAPBUSINESS.reg" -Wait
            Start-Sleep 20
            Write-Host -ForegroundColor Green " DONE"

            $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
            "$date Installazione SAP" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
        }
    }
    if ($opt_sw_string -match 'SAP Analysis') {
        $upgrade_sap = 'yes'
    }
    if ($upgrade_sap -eq 'yes') {
        $source_path = '"\\itmilitgroup\SD_Utilities\UtilitiesSD\SAP Analysis_Office365"'
        $dest_path = '"C:\TEMPSOFTWARE\SAP Analysis_Office365"'
        Write-Host -NoNewline "Recupero SAP Analysis"
        $string = "$source_path $dest_path /E /Z"
        Start-Process -Wait robocopy $string
        Write-Host -ForegroundColor Green " DONE"

        Write-Host -NoNewline "Installo SAP Analysis"
        foreach ($exec in ('vstor_redist.exe', 'AOFFICE28SP00_0-70004973.EXE', 'AOFFICE28SP01_0-70004973.EXE')) {
            $string = "C:\TEMPSOFTWARE\SAP Analysis_Office365\$exec"
            Start-Process -Wait $string
        }   
        Write-Host -ForegroundColor Green " DONE"

        [System.Windows.MessageBox]::Show('Lanciare un app chiamata "Analysis for Microsoft Excel"','APP','Ok','Info') > $null

        $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
        "$date Installazione SAP Analysis" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
    }

}

if ($optional_usb.Checked -eq $true) {
    Write-Host -ForegroundColor Yellow "`n*** Installazione periferiche USB ***"

    $optusb_sourcepath = $prefix + "\TEMPSOFTWARE\usb_devices.log"
    Copy-Item $optusb_sourcepath -Destination "C:\TEMPSOFTWARE" > $null

    $ErrorActionPreference = 'SilentlyContinue'
    $opt_usb_list = Get-Content -Path "C:\TEMPSOFTWARE\usb_devices.log"
    $opt_usb_string = [system.String]::Join(";", $opt_usb_list)
    $ErrorActionPreference= 'Inquire'

    $opt_list = @{
        HP_DeskJet_1010 = @{
            string = 'HP Deskjet 1010 series'
            source_path = '\\itmilitgroup\SD_Utilities\UtilitiesSD\DRIVERS\DriverStampanti\Driver stampanti win 10\DJ1010_188.exe'
            dest_path = 'C:\TEMPSOFTWARE\DJ1010_188.exe'
            exe_path = 'C:\TEMPSOFTWARE\DJ1010_188.exe'
        }
        HP_DeskJet_1110b = @{
            string = 'HP DJ 1110 series'
            source_path = '\\itmilitgroup\SD_Utilities\UtilitiesSD\DRIVERS\DriverStampanti\Driver stampanti win 10\DJ1110_Full_WebPack_40.11.1124.exe'
            dest_path = 'C:\TEMPSOFTWARE\DJ1110_Full_WebPack_40.11.1124.exe'
            exe_path = 'C:\TEMPSOFTWARE\DJ1110_Full_WebPack_40.11.1124.exe'
        }
        HP_DeskJet_1110 = @{
            string = 'HP DeskJet 1110 series'
            source_path = '\\itmilitgroup\SD_Utilities\UtilitiesSD\DRIVERS\DriverStampanti\Driver stampanti win 10\DJ1110_Full_WebPack_40.11.1124.exe'
            dest_path = 'C:\TEMPSOFTWARE\DJ1110_Full_WebPack_40.11.1124.exe'
            exe_path = 'C:\TEMPSOFTWARE\DJ1110_Full_WebPack_40.11.1124.exe'
        }
        HP_DeskJet_3700 = @{
            string = 'HP DeskJet 3700 series'
            source_path = '\\itmilitgroup\SD_Utilities\UtilitiesSD\DRIVERS\DriverStampanti\Driver stampanti win 10\DJ3700_Basicx64_40.12.1161.exe'
            dest_path = 'C:\TEMPSOFTWARE\DJ3700_Basicx64_40.12.1161.exe'
            exe_path = 'C:\TEMPSOFTWARE\DJ3700_Basicx64_40.12.1161.exe'
        }
        ScanSnap_1500 = @{
            string = 'Image ScanSnap S1500'
            source_path = '\\itmilitgroup\SD_Utilities\UtilitiesSD\DRIVERS\Driver Scanner\Driver Scanner_1500'
            dest_path = 'C:\TEMPSOFTWARE\Driver Scanner_1500'
            exe_path = 'C:\TEMPSOFTWARE\Driver Scanner_1500\setup.exe'
        }
        ScanSnap_IX1500 = @{
            string = 'Image ScanSnap iX500'
            source_path = '\\itmilitgroup\SD_Utilities\UtilitiesSD\DRIVERS\Driver Scanner\Scanner IX500'
            dest_path = 'C:\TEMPSOFTWARE\Scanner IX500'
            exe_path = 'C:\TEMPSOFTWARE\Scanner IX500\WiniX500ManagerV65L61WW.exe'
        }
    }

    foreach ($key in $opt_list.Keys) {
        if ($opt_usb_string -match $opt_list[$key].string) {
            [System.Windows.MessageBox]::Show("Installazione $key",'CONFIGURAZIONE SW','Ok','Info') > $null

            Write-Host -NoNewline "Recupero" $key
            Copy-Item -Recurse $opt_list[$key].source_path -Destination $opt_list[$key].dest_path > $null
            Write-Host -ForegroundColor Green " DONE"
        
            Write-Host -NoNewline "Installo" $key
            Start-Process $opt_list[$key].exe_path
            Start-Sleep 20
            Write-Host -ForegroundColor Green " DONE"

            $date =  Get-Date -Format "[yyyy/MM/dd - hh:mm:ss]"
            "$date Installazione $key" | Out-File "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Encoding ASCII -Append
        }
    }
}

if ($sendlog.Checked -eq $true) {
    if (Test-Path "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -PathType Leaf) {
        # mando sul ticket una mail con il file di log
        $form_EXP = FormBase -w 300 -h 150 -text "EXPERIMENTAL"
        $label = New-Object System.Windows.Forms.Label
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(150,30)
        $label.Text = "Inserire il ticket ID:"
        $form_EXP.Controls.Add($label)
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(160,20)
        $textBox.Size = New-Object System.Drawing.Size(120,30)
        $form_EXP.Controls.Add($textBox)
        OKButton -form $form_EXP -x 100 -y 50 -text "Ok"
        $form_EXP.Add_Shown({$textBox.Select()})
        $result = $form_EXP.ShowDialog()
        $header = 'RE: Richiesta di assistenza - #' + $textBox.Text
    
        $Outlook = New-Object -ComObject Outlook.Application
        $Mail = $Outlook.CreateItem(0)
        $Mail.To = "it-fm-itsm@kpmg.it" 
        $Mail.Subject = $header
        $Mail.Body = "LOG della migrazione dati sul PC $env:COMPUTERNAME"
        $Mail.Attachments.Add("C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log");
        $Mail.Send()
        $Outlook.Quit() 
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null
    }
}

# copia del file di log
if ($itmilfsrsd_logged -eq $true) {
    $ErrorActionPreference = 'Stop'
    Try {
        $alertpath = "R:\$env:USERNAME-$env:COMPUTERNAME"
        if (!(Test-Path $alertpath)) {
            New-Item -ItemType directory -Path $alertpath > $null
        }
        Copy-Item "C:\TEMPSOFTWARE\MIGRAZIONE-DATI.log" -Destination $alertpath > $null
    }
    Catch {
        Write-Host -ForegroundColor Red "`nImpossibile copiare file di log"
    }
    $ErrorActionPreference = 'Inquire'
}

if (!($modalita -eq 'null')) {
    Remove-PSDrive -Name R
}

# pulizia temporanei
[System.Windows.MessageBox]::Show("Clicca OK quando le installazioni saranno ultimate",'CONFIGURAZIONE SW','Ok','Warning') > $null
Remove-Item "C:\TEMPSOFTWARE" -Recurse -Force

<#
# registrazione
$answ = [System.Windows.MessageBox]::Show("Registrare il cambio PC?",'REGISTRAZIONE','YesNo','Info')
if ($answ -eq "Yes") {    
    Import-Module -Name '\\itmilitgroup\SD_Utilities\SCRIPT\Moduli_PowerShell\Runas.psm1'
    Runa -executable '\\itmilitgroup\SD_Utilities\SCRIPT\CAMBIO_PC\CAMBIO_PC.exe' -ITuser $login.UserName
}
#>

# reboot
$answ = [System.Windows.MessageBox]::Show("Riavvio computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {
    Restart-Computer
}
