<#
Name......: Inizializza_PC.ps1
Version...: 20.12.1
Author....: Dario CORRADA

Questo script rinomina un PC usando il suo seriale, lo mette a dominio e crea un'utenza locale con privilegi di admin
#>

# header
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

Import-Module -Name '\\192.168.2.251\Dario\SCRIPT\Moduli_PowerShell\Forms.psm1'

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

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

Remove-Item $tmppath -Recurse -Force


# creo utenza locale
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


# metto a dominio il PC
$dominio = 'agm.local'

$form_PWD = FormBase -w 400 -h 250 -text "LOGIN"
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Size(10,20) 
$label.Size = New-Object System.Drawing.Size(300,20) 
$label.Text = "Inserisci le tue credenziali"
$form_PWD.Controls.Add($label)
$usrlabel = New-Object System.Windows.Forms.Label
$usrlabel.Location = New-Object System.Drawing.Size(10,50) 
$usrlabel.Size = New-Object System.Drawing.Size(100,20) 
$usrlabel.Text = "Utente:"
$form_PWD.Controls.Add($usrlabel)
$pwdlabel = New-Object System.Windows.Forms.Label
$pwdlabel.Location = New-Object System.Drawing.Size(10,80) 
$pwdlabel.Size = New-Object System.Drawing.Size(100,20) 
$pwdlabel.Text = "Password:"
$form_PWD.Controls.Add($pwdlabel)
$MaskedTextBox = New-Object System.Windows.Forms.MaskedTextBox
$MaskedTextBox.PasswordChar = '*'
$MaskedTextBox.Location = New-Object System.Drawing.Point(130,80)
$MaskedTextBox.Size = New-Object System.Drawing.Size(150,20)
$form_PWD.Add_Shown({$MaskedTextBox.Select()})
$form_PWD.Controls.Add($MaskedTextBox)
$textBox = New-Object System.Windows.Forms.TextBox
$textBox.Location = New-Object System.Drawing.Point(130,50)
$textBox.Size = New-Object System.Drawing.Size(150,20)
$form_PWD.Add_Shown({$textBox.Select()})
$form_PWD.Controls.Add($textBox)
OKButton -form $form_PWD -x 100 -y 120 -text "Ok"
$result = $form_PWD.ShowDialog()
$usr = $textBox.Text
$pwd = ConvertTo-SecureString $MaskedTextBox.Text -AsPlainText -Force
$ad_login = New-Object System.Management.Automation.PSCredential($usr, $pwd)

# form di scelta OU
$form_modalita = FormBase -w 400 -h 230 -text "OU DESTINAZIONE"
$noou = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "null"
$interni  = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "utenti_interni"
$esterni = RadioButton -form $form_modalita -checked $false -x 30 -y 80 -text "utenti_esterni"
OKButton -form $form_modalita -x 90 -y 130 -text "Ok"
$result = $form_modalita.ShowDialog()

if ($result -eq "OK") {
    if ($noou.Checked) {
        $outarget = "null"
    } elseif ($interni.Checked) {
        $outarget = 'OU=utenti_interni,DC=agm,DC=local'
    } elseif ($esterni.Checked) {
        $outarget = 'OU=utenti_esterni,DC=agm,DC=local'
    }    
}

Write-Host "Domain...: " -NoNewline
Write-Host $dominio -ForegroundColor Cyan
Write-Host "OU.......: " -NoNewline
Write-Host $outarget -ForegroundColor Cyan

$ErrorActionPreference= 'Stop'
Try {
    if ($outarget -eq "null") {
        Add-Computer -ComputerName $hostname -Credential $ad_login -DomainName $dominio -Force
    } else {
        Add-Computer -ComputerName $hostname -Credential $ad_login -DomainName $dominio -OUPath $outarget -Force
    }
    Write-Host "PC messo a dominio" -ForegroundColor Green
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



