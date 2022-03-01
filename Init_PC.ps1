<#
Name......: Init_PC.ps1
Version...: 20.12.3
Author....: Dario CORRADA

This script finalize fresh OS installations:
* install Chrome, AcrobatDC, 7Zip, Supremo;
* create a local account with admin privileges;
* set hostname according to the serial number.
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Init_PC\.ps1$" > $null
$workdir = $matches[1]

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# fetch and install additional softwares
# modify download paths according to updated software versions (updated on 2021/01/18)
$tmppath = "C:\TEMPSOFTWARE"
New-Item -ItemType directory -Path $tmppath > $null
$swlist = @{}
$form_panel = FormBase -w 350 -h 425 -text "SOFTWARES"
$swlist['Acrobat Reader DC'] = CheckBox -form $form_panel -checked $true -x 20 -y 20 -text "Acrobat Reader DC"
$swlist['Chrome'] = CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "Chrome"
$swlist['Revo Uninstaller'] = CheckBox -form $form_panel -checked $true -x 20 -y 80 -text "Revo Uninstaller"
$swlist['Skype'] = CheckBox -form $form_panel -checked $true -x 20 -y 110 -text "Skype"
$swlist['Skype'].Checked = $false
$swlist['Speccy'] = CheckBox -form $form_panel -checked $true -x 20 -y 140 -text "Speccy"
$swlist['Supremo'] = CheckBox -form $form_panel -checked $true -x 20 -y 170 -text "Supremo"
$swlist['Teams'] = CheckBox -form $form_panel -checked $true -x 20 -y 200 -text "Teams"
$swlist['WinDirStat'] = CheckBox -form $form_panel -checked $true -x 20 -y 230 -text "WinDirStat"
$swlist['7ZIP'] = CheckBox -form $form_panel -checked $true -x 20 -y 260 -text "7ZIP"
OKButton -form $form_panel -x 100 -y 310 -text "Ok"
$result = $form_panel.ShowDialog()

$download = New-Object net.webclient
foreach ($item in ($swlist.Keys | Sort-Object)) {
    if ($swlist[$item].Checked -eq $true) {
        Write-Host -ForegroundColor Blue "[$item]"
        if ($item -eq 'Acrobat Reader DC') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("http://ardownload.adobe.com/pub/adobe/reader/win/AcrobatDC/1900820071/AcroRdrDC1900820071_it_IT.exe", "$tmppath\AcroReadDC.exe")
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Start-Process -FilePath "$tmppath\AcroReadDC.exe" -Wait
            Write-Host -ForegroundColor Green " DONE"     
        } elseif ($item -eq 'Chrome') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("http://dl.google.com/chrome/install/375.126/chrome_installer.exe", "$tmppath\ChromeSetup.exe")
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Start-Process -FilePath "$tmppath\ChromeSetup.exe" -Wait
            Write-Host -ForegroundColor Green " DONE"   
        } elseif ($item -eq 'Revo Uninstaller') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://www.revouninstaller.com/download-freeware-version.php", "$tmppath\Revo.exe")
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Start-Process -FilePath "$tmppath\Revo.exe" -Wait
            Write-Host -ForegroundColor Green " DONE"   
        } elseif ($item -eq 'Skype') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://go.skype.com/windows.desktop.download", "$tmppath\Skype.exe")
            Write-Host -ForegroundColor Green " DONE"
            $answ = [System.Windows.MessageBox]::Show("After installation close all Skype instances to proceed...",'WARNING','Ok','Warning')
            Write-Host -NoNewline "Install software..."
            Start-Process -FilePath "$tmppath\Skype.exe" -Wait
            Write-Host -ForegroundColor Green " DONE"        
        } elseif ($item -eq 'Speccy') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://download.ccleaner.com/spsetup132.exe", "$tmppath\Speccy.exe")
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Start-Process -FilePath "$tmppath\Speccy.exe" -Wait
            Write-Host -ForegroundColor Green " DONE"   
        } elseif ($item -eq 'Supremo') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://www.agmsolutions.net/wp-content/uploads/assistenza/Assistenza_Remota.exe", "$env:PUBLIC\Desktop\Supremo.exe")
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Write-Host -ForegroundColor Green " DONE"
        } elseif ($item -eq 'Teams') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://go.microsoft.com/fwlink/p/?LinkID=869426&clcid=0x410&culture=it-it&country=IT&lm=deeplink&lmsrc=groupChatMarketingPageWeb&cmpid=directDownloadWin64", "C:\Users\Public\Desktop\Teams_Installer.exe")
            Write-Host -ForegroundColor Green " DONE"
            $answ = [System.Windows.MessageBox]::Show("Please run setup once the target account has been logged in",'INFO','Ok','Info')
        } elseif ($item -eq 'WinDirStat') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://windirstat.mirror.wearetriple.com//wds_current_setup.exe", "$tmppath\WinDirStat.exe")
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Start-Process -FilePath "$tmppath\WinDirStat.exe" -Wait
            Write-Host -ForegroundColor Green " DONE"   
        } elseif ($item -eq '7ZIP') {
            Write-Host -NoNewline "Download software..."
            $download.Downloadfile("https://www.7-zip.org/a/7z1900-x64.exe", "$tmppath\7Zip.exe")
            Write-Host -ForegroundColor Green " DONE"
            Write-Host -NoNewline "Install software..."
            Start-Process -FilePath "$tmppath\7Zip.exe" -Wait
            Write-Host -ForegroundColor Green " DONE"   
        }
    }
}

# remove Skype startup
New-PSDrive -PSProvider Registry -Name HKU -Root HKEY_USERS | Out-Null
$startups = Get-CimInstance Win32_StartupCommand | Select-Object Name,Location
foreach ($startup in $startups){
    if ($startup.Name -eq 'Skype for Desktop'){
        $number = ($startup.Location).IndexOf("\")
        $location = ($startup.Location).Insert("$number",":")
        Write-Output "Disabling $($startup.Name) from $location)"
        Remove-ItemProperty -Path "$location" -Name "$($startup.name)" 
    }
}

Remove-Item $tmppath -Recurse -Force


# creating local account
$answ = [System.Windows.MessageBox]::Show("Create local account?",'ACCOUNT','YesNo','Info')
if ($answ -eq "Yes") {
    
    $form = FormBase -w 520 -h 270 -text "ACCOUNT"
    $font = New-Object System.Drawing.Font("Arial", 12)
    $form.Font = $font

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(500,30)
    $label.Text = "Username:"
    $form.Controls.Add($label)

    $usrname = New-Object System.Windows.Forms.TextBox
    $usrname.Location = New-Object System.Drawing.Point(10,60)
    $usrname.Size = New-Object System.Drawing.Size(450,30)
    $form.Controls.Add($usrname)

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,100)
    $label2.Size = New-Object System.Drawing.Size(500,30)
    $label2.Text = "Fullname:"
    $form.Controls.Add($label2)

    $fullname = New-Object System.Windows.Forms.TextBox
    $fullname.Location = New-Object System.Drawing.Point(10,140)
    $fullname.Size = New-Object System.Drawing.Size(450,30)
    $form.Controls.Add($fullname)

    $OKButton = New-Object System.Windows.Forms.Button

    OKButton -form $form -x 200 -y 190 -text "Ok"

    $form.Topmost = $true
    $result = $form.ShowDialog()

    $username = $usrname.Text
    $completo = $fullname.Text

    $form_pswd = FormBase -w 450 -h 230 -text "CREATE PASSWORD"
    $personal = RadioButton -form $form_pswd -checked $true -x 30 -y 20 -text "Set your own password"
    $randomic  = RadioButton -form $form_pswd -checked $false -x 30 -y 50 -text "Generate random password"
    OKButton -form $form_pswd -x 90 -y 120 -text "Ok"
    $result = $form_pswd.ShowDialog()
    if ($result -eq "OK") {
        if ($personal.Checked) {
            $form = FormBase -w 520 -h 200 -text "PASSWORD"
            $font = New-Object System.Drawing.Font("Arial", 12)
            $form.Font = $font
        
            $label = New-Object System.Windows.Forms.Label
            $label.Location = New-Object System.Drawing.Point(10,20)
            $label.Size = New-Object System.Drawing.Size(500,30)
            $label.Text = "Password:"
            $form.Controls.Add($label)
        
            $usrname = New-Object System.Windows.Forms.TextBox
            $usrname.Location = New-Object System.Drawing.Point(10,60)
            $usrname.Size = New-Object System.Drawing.Size(450,30)
            $usrname.PasswordChar = '*'
            $form.Controls.Add($usrname)
        
            $OKButton = New-Object System.Windows.Forms.Button
        
            OKButton -form $form -x 200 -y 120 -text "Ok"
        
            $form.Topmost = $true
            $result = $form.ShowDialog()

            $thepasswd = $usrname.Text
        } elseif ($randomic.Checked) {
            Add-Type -AssemblyName 'System.Web'
            $thepasswd = [System.Web.Security.Membership]::GeneratePassword(10, 0)
        }
    }
    Write-Host "Username...: " -NoNewline
    Write-Host $username -ForegroundColor Cyan
    Write-Host "Password...: " -NoNewline
    Write-Host $thepasswd -ForegroundColor Cyan
    Pause
    $pwd = ConvertTo-SecureString $thepasswd -AsPlainText -Force
    
    $ErrorActionPreference= 'Stop'
    Try {
        New-LocalUser -Name $username -Password $pwd -FullName $completo -PasswordNeverExpires -AccountNeverExpires -Description "utente locale"
        Add-LocalGroupMember -Group "Administrators" -Member $username
        Write-Host -ForegroundColor Green "Local account created"
        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
        exit
    }    
}

# changing hostname
$serial = Get-WmiObject win32_bios
$hostname = $serial.SerialNumber
$answ = [System.Windows.MessageBox]::Show("Il PC e' dislocato a Torino?",'LOCAZIONE','YesNo','Info')
if ($answ -eq "Yes") {    
    $hostname = $hostname + '-TO'
}
Write-Host "Hostname...: " -NoNewline
Write-Host $hostname -ForegroundColor Cyan

$ErrorActionPreference= 'Stop'
Try {
    Rename-Computer -NewName $hostname
    Write-Host "PC renamed" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}  

# reboot
$answ = [System.Windows.MessageBox]::Show("Reboot computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}
