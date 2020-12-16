<#
Name......: UpdateDB.ps1
Version...: 20.1.1
Author....: Dario CORRADA

Questo script gestice il DB di Patrol

+++ UPDATES +++

[2019-10-08  CORRADA] 
Vedi GIT
#>

# header
$ErrorActionPreference= 'Inquire'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

Import-Module -Name '\\192.168.2.251\Dario\SCRIPT\Moduli_PowerShell\Forms.psm1'
Import-Module -Name '\\192.168.2.251\Dario\SCRIPT\Moduli_PowerShell\Patrol.psm1'
Import-Module -Name '\\192.168.2.251\Dario\SCRIPT\Moduli_PowerShell\FileCryptography.psm1'

function RetryButton {
    param ($form, $x, $y, $text)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = "$x,$y"
    $OKButton.Size = '100,30'
    $OKButton.Text = $text
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::Retry
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)
}

$logo = Patrol -scriptname UpdateDB

# form di scelta modalita'
do {
    $form_modalita = FormBase -w 350 -h 300 -text "MODALITA'"
    $importa = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "Importa Database"
    $esporta = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "Esporta Database"
    $add_user = RadioButton -form $form_modalita -checked $false -x 30 -y 80 -text "Aggiungi record"
    $check_user = RadioButton -form $form_modalita -checked $false -x 30 -y 110 -text "Verifica privilegi utente"
    $check_script = RadioButton -form $form_modalita -checked $false -x 30 -y 140 -text "Verifica accessi script"
    OKButton -form $form_modalita -x 190 -y 180 -text "Exit"
    RetryButton -form $form_modalita -x 50 -y 180 -text "Next"
    $ciclo = $form_modalita.ShowDialog()

    if ($ciclo -eq "Retry") {
        if ($add_user.Checked) {
            $form_add = FormBase -w 400 -h 200 -text "AGGIUNGI RECORD:"
            $usrlabel = New-Object System.Windows.Forms.Label
            $usrlabel.Location = New-Object System.Drawing.Size(10,30) 
            $usrlabel.Size = New-Object System.Drawing.Size(100,20) 
            $usrlabel.Text = "Utente:"
            $form_add.Controls.Add($usrlabel)
            $usrBox = New-Object System.Windows.Forms.TextBox
            $usrBox.Location = New-Object System.Drawing.Point(130,30)
            $usrBox.Size = New-Object System.Drawing.Size(150,20)
            $form_add.Add_Shown({$usrBox.Select()})
            $form_add.Controls.Add($usrBox)
            $scriptlabel = New-Object System.Windows.Forms.Label
            $scriptlabel.Location = New-Object System.Drawing.Size(10,60) 
            $scriptlabel.Size = New-Object System.Drawing.Size(100,20) 
            $scriptlabel.Text = "Script:"
            $form_add.Controls.Add($scriptlabel)
            $scriptBox = New-Object System.Windows.Forms.TextBox
            $scriptBox.Location = New-Object System.Drawing.Point(130,60)
            $scriptBox.Size = New-Object System.Drawing.Size(150,20)
            $form_add.Add_Shown({$scriptBox.Select()})
            $form_add.Controls.Add($scriptBox)
            OKButton -form $form_add -x 100 -y 120 -text "Ok"
            $result = $form_add.ShowDialog()

            $string = $scriptBox.Text + ';' + $usrBox.Text

            Copy-Item -Path "\\192.168.2.251\Dario\SCRIPT\PATROL\PatrolDB.csv.AES" -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
            $stringa = Get-Content '\\192.168.2.251\Dario\SCRIPT\PATROL\crypto.key'
            $key = ConvertTo-SecureString $stringa -AsPlainText -Force
            Unprotect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Algorithm AES -Key $key -RemoveSource

            $string | Out-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv" -Encoding ASCII -Append

            net stop workstation /y > $null
            net start workstation > $null
            Start-Sleep 5
            $ErrorActionPreference = 'Stop'
            Try {
                New-PSDrive -Name P -PSProvider FileSystem -Root "\\192.168.2.251\Dario\SCRIPT\PATROL" -Credential $logo > $null
            }
            Catch { 
                $errormsg = [System.Windows.MessageBox]::Show("Impossibile accedere alla cartella remota, controllare le credenziali",'ATTENZIONE','Ok','Error')
            }
            $ErrorActionPreference = 'Inquire'
    
            $stringa = Get-Content 'P:\crypto.key'
            $key = ConvertTo-SecureString $stringa -AsPlainText -Force
            Protect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv" -Algorithm AES -Key $key -RemoveSource
            Remove-Item -Path "P:\PatrolDB.csv.AES"
            Copy-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Destination "P:\PatrolDB.csv.AES"
            Remove-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
    
            # mi disconnetto da \\itmilitgroup
            Remove-PSDrive -Name P

        }

        if ($check_user.Checked) {
            Copy-Item -Path "\\192.168.2.251\Dario\SCRIPT\PATROL\PatrolDB.csv.AES" -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
            $stringa = Get-Content '\\192.168.2.251\Dario\SCRIPT\PATROL\crypto.key'
            $key = ConvertTo-SecureString $stringa -AsPlainText -Force
            Unprotect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Algorithm AES -Key $key -RemoveSource

            $filecontent = Get-Content -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv"
            $lista_utenti = @()
            foreach ($newline in $filecontent) {
                ($script, $utente) = $newline.Split(';')
                if (!($lista_utenti -contains $utente)) {
                    if (!($utente -eq 'USER')) {
                        $lista_utenti += $utente
                    }
                }
            }

            $formlist = FormBase -w 400 -h 200 -text "LISTA UTENTI"
            $DropDown = new-object System.Windows.Forms.ComboBox
            $DropDown.Location = new-object System.Drawing.Size(10,60)
            $DropDown.Size = new-object System.Drawing.Size(250,30)
            foreach ($elem in ($lista_utenti | sort)) {
                $DropDown.Items.Add($elem)  > $null
            }
            $formlist.Controls.Add($DropDown)
            $DropDownLabel = new-object System.Windows.Forms.Label
            $DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
            $DropDownLabel.size = new-object System.Drawing.Size(500,30) 
            $DropDownLabel.Text = "Scegliere l'utente"
            $formlist.Controls.Add($DropDownLabel)
            OKButton -form $formlist -x 100 -y 100 -text "Ok"
            $formlist.Add_Shown({$DropDown.Select()})
            $result = $formlist.ShowDialog()
            $selected = $DropDown.Text
    
            $lista_risultato = @()
            foreach ($newline in $filecontent) {
                ($script, $utente) = $newline.Split(';')
                if ($utente -eq $selected) {
                    $lista_risultato += $script
                }
            }
            Write-Host -ForegroundColor Cyan "Script a cui e' abilitato $selected"
            $lista_risultato | sort
                
            Remove-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv"
        }

        if ($check_script.Checked) {
            Copy-Item -Path "\\192.168.2.251\Dario\SCRIPT\PATROL\PatrolDB.csv.AES" -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
            $stringa = Get-Content '\\192.168.2.251\Dario\SCRIPT\PATROL\crypto.key'
            $key = ConvertTo-SecureString $stringa -AsPlainText -Force
            Unprotect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Algorithm AES -Key $key -RemoveSource

            $filecontent = Get-Content -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv"
            $lista_script = @()
            foreach ($newline in $filecontent) {
                ($script, $utente) = $newline.Split(';')
                if (!($lista_script -contains $script)) {
                    if (!($script -eq 'SCRIPT')) {
                        $lista_script += $script
                    }
                }
            }

            $formlist = FormBase -w 400 -h 200 -text "LISTA SCRIPT"
            $DropDown = new-object System.Windows.Forms.ComboBox
            $DropDown.Location = new-object System.Drawing.Size(10,60)
            $DropDown.Size = new-object System.Drawing.Size(250,30)
            foreach ($elem in ($lista_script | sort)) {
                $DropDown.Items.Add($elem)  > $null
            }
            $formlist.Controls.Add($DropDown)
            $DropDownLabel = new-object System.Windows.Forms.Label
            $DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
            $DropDownLabel.size = new-object System.Drawing.Size(500,30) 
            $DropDownLabel.Text = "Scegliere lo script"
            $formlist.Controls.Add($DropDownLabel)
            OKButton -form $formlist -x 100 -y 100 -text "Ok"
            $formlist.Add_Shown({$DropDown.Select()})
            $result = $formlist.ShowDialog()
            $selected = $DropDown.Text
    
            $lista_risultato = @()
            foreach ($newline in $filecontent) {
                ($script, $utente) = $newline.Split(';')
                if ($script -eq $selected) {
                    $lista_risultato += $utente
                }
            }
            Write-Host -ForegroundColor Cyan "Utenti che hanno accesso a $selected"
            $lista_risultato | sort

            Remove-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv"
        }

        if ($importa.Checked) {
            Write-Host "Importo il database..."

            net stop workstation /y > $null
            net start workstation > $null
            Start-Sleep 5
            $ErrorActionPreference = 'Stop'
            Try {
                New-PSDrive -Name P -PSProvider FileSystem -Root "\\192.168.2.251\Dario\SCRIPT\PATROL" -Credential $logo > $null
            }
            Catch { 
                $errormsg = [System.Windows.MessageBox]::Show("Impossibile accedere alla cartella remota, controllare le credenziali",'ATTENZIONE','Ok','Error')
            }
            $ErrorActionPreference = 'Inquire'
            
            # finestra di dialogo per selezionare il file CSV di input di Varonis
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
            $OpenFileDialog.filter = "Comma separated value (*.csv)| *.csv"
            $OpenFileDialog.ShowDialog() | Out-Null
            $infile = $OpenFileDialog.filename
    
            Copy-Item -Path $infile -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv"
            $stringa = Get-Content 'P:\crypto.key'
            $key = ConvertTo-SecureString $stringa -AsPlainText -Force
            Protect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv" -Algorithm AES -Key $key -RemoveSource
            Remove-Item -Path "P:\PatrolDB.csv.AES"
            Copy-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Destination "P:\PatrolDB.csv.AES"
            Remove-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
            [System.Windows.MessageBox]::Show("DB importato come \\192.168.2.251\Dario\SCRIPT\PATROL\PatrolDB.csv.AES",'PATROL DB','Ok','Info') > $null
    
            # mi disconnetto da \\itmilitgroup
            Remove-PSDrive -Name P
        }
    
        if ($esporta.Checked) {
            Write-Host "Esporto il database..."
            Copy-Item -Path "\\192.168.2.251\Dario\SCRIPT\PATROL\PatrolDB.csv.AES" -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
            $stringa = Get-Content '\\192.168.2.251\Dario\SCRIPT\PATROL\crypto.key'
            $key = ConvertTo-SecureString $stringa -AsPlainText -Force
            Unprotect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Algorithm AES -Key $key -RemoveSource
    
            # finestra di dialogo per salvare il file CSV di output
            $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $SaveFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
            $SaveFileDialog.filter = "Comma separated value (*.csv)| *.csv"
            $SaveFileDialog.filename = $filename
            $SaveFileDialog.ShowDialog() | Out-Null
            $outfile = $SaveFileDialog.filename
            Move-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv" -Destination $outfile
    
            [System.Windows.MessageBox]::Show("DB esportato come $outfile",'PATROL DB','Ok','Info') > $null
        }
    }
} until ($ciclo -eq "Ok")



