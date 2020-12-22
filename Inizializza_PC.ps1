<#
Name......: Inizializza_PC.ps1
Version...: 20.12.3
Author....: Dario CORRADA

Questo script inizializza un PC da assegnare:
* installa Chrome, AcrobatDC e 7Zip;
* crea un'utenza locale con privilegi di admin;
* rinomina il PC usando il suo seriale;
#>

# faccio in modo di elevare l'esecuzione dello script con privilegi di admin
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# recupero il percorso di installazione
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Certificati\.ps1$" > $null
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
$workdir = Get-Location
Import-Module -Name "$workdir\Moduli_PowerShell\Forms.psm1"

# download e installazione software
# modificare i path dei download per scaricare software aggiornato
$tmppath = "C:\TEMPSOFTWARE"
New-Item -ItemType directory -Path $tmppath > $null
Write-Host -NoNewline "Download software..."
$download = New-Object net.webclient
$download.Downloadfile("http://dl.google.com/chrome/install/375.126/chrome_installer.exe", "$tmppath\ChromeSetup.exe")
$download.Downloadfile("http://ardownload.adobe.com/pub/adobe/reader/win/AcrobatDC/1502320053/AcroRdrDC1502320053_en_US.exe", "$tmppath\AcroReadDC.exe")
$download.Downloadfile("https://www.7-zip.org/a/7z1900-x64.exe", "$tmppath\7Zip.exe")
Write-Host -ForegroundColor Green " DONE"

Write-Host -NoNewline "Installazione software..."
Start-Process -FilePath "$tmppath\ChromeSetup.exe" -Wait
Start-Process -FilePath "$tmppath\AcroReadDC.exe" -Wait
Start-Process -FilePath "$tmppath\7Zip.exe" -Wait
Write-Host -ForegroundColor Green " DONE"

Remove-Item $tmppath -Recurse -Force


# creazione utenza locale
$answ = [System.Windows.MessageBox]::Show("Creare utenza locale?",'ACCOUNT','YesNo','Info')
if ($answ -eq "Yes") {
    
    $form = FormBase -w 520 -h 220 -text "ACCOUNT"
    $font = New-Object System.Drawing.Font("Arial", 12)
    $form.Font = $font

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(500,30)
    $label.Text = "Nome utente:"
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,60)
    $textBox.Size = New-Object System.Drawing.Size(450,30)
    $form.Controls.Add($textBox)
    $OKButton = New-Object System.Windows.Forms.Button

    OKButton -form $form -x 200 -y 110 -text "Ok"

    $form.Topmost = $true
    $result = $form.ShowDialog()

    $username = $textBox.Text
    $passwd = "Password1"
    Write-Host "Username...: " -NoNewline
    Write-Host $username -ForegroundColor Cyan
    Write-Host "Password...: " -NoNewline
    Write-Host $passwd -ForegroundColor Cyan
    $pwd = ConvertTo-SecureString $passwd -AsPlainText -Force
    
    $ErrorActionPreference= 'Stop'
    Try {
        New-LocalUser -Name $username -Password $pwd -PasswordNeverExpires -AccountNeverExpires -Description "utente locale"
        Add-LocalGroupMember -Group "Administrators" -Member $username
        Write-Host -ForegroundColor Green "Account locale creato"
        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
        exit
    }    
}

# rinomino il PC
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
    Write-Host "PC rinominato" -ForegroundColor Green
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}  

# reboot
$answ = [System.Windows.MessageBox]::Show("Riavvio computer?",'REBOOT','YesNo','Info')
if ($answ -eq "Yes") {    
    Restart-Computer
}



