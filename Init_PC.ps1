<#
Name......: Init_PC.ps1
Version...: 20.12.3
Author....: Dario CORRADA

Questo script finalize fresh OS installations:
* install Chrome, AcrobatDC, 7Zip;
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
Write-Host -NoNewline "Download software..."
$download = New-Object net.webclient
$download.Downloadfile("http://dl.google.com/chrome/install/375.126/chrome_installer.exe", "$tmppath\ChromeSetup.exe")
$download.Downloadfile("http://ardownload.adobe.com/pub/adobe/reader/win/AcrobatDC/1900820071/AcroRdrDC1900820071_it_IT.exe", "$tmppath\AcroReadDC.exe")
$download.Downloadfile("https://www.7-zip.org/a/7z1900-x64.exe", "$tmppath\7Zip.exe")
$download.Downloadfile("https://www.supremocontrol.com/download.aspx?file=Supremo.exe&id_sw=7&ws=supremocontrol.com", "$env:PUBLIC\Desktop\Supremo.exe")
Write-Host -ForegroundColor Green " DONE"

Write-Host -NoNewline "Install software..."
Start-Process -FilePath "$tmppath\ChromeSetup.exe" -Wait
Start-Process -FilePath "$tmppath\AcroReadDC.exe" -Wait
Start-Process -FilePath "$tmppath\7Zip.exe" -Wait
Write-Host -ForegroundColor Green " DONE"

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
    $passwd = "Password1"
    Write-Host "Username...: " -NoNewline
    Write-Host $username -ForegroundColor Cyan
    Write-Host "Password...: " -NoNewline
    Write-Host $passwd -ForegroundColor Cyan
    $pwd = ConvertTo-SecureString $passwd -AsPlainText -Force
    
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
