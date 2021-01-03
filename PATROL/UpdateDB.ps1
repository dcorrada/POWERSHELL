<#
Name......: UpdateDB.ps1
Version...: 20.1.1
Author....: Dario CORRADA

Questo script manage Patrol DB
#>

# elevate script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# retrieve installation path
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\UpdateDB\.ps1$" > $null
$repopath = $matches[1]

# setting execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# import modules
Import-Module -Name "$repopath\Modules\Forms.psm1"
Import-Module -Name "$repopath\Modules\Patrol.psm1"
Import-Module -Name "$repopath\Modules\FileCryptography.psm1"

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

# dialog box
do {
    $form_modalita = FormBase -w 350 -h 300 -text "OPTIONS'"
    $importa = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "Import Database"
    $esporta = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "Export Database"
    $add_user = RadioButton -form $form_modalita -checked $false -x 30 -y 80 -text "Add record"
    $check_user = RadioButton -form $form_modalita -checked $false -x 30 -y 110 -text "Check user privileges"
    $check_script = RadioButton -form $form_modalita -checked $false -x 30 -y 140 -text "Check script grants"
    OKButton -form $form_modalita -x 190 -y 180 -text "Exit"
    RetryButton -form $form_modalita -x 50 -y 180 -text "Next"
    $ciclo = $form_modalita.ShowDialog()

    if ($ciclo -eq "Retry") {
        if ($add_user.Checked) {
            $form_add = FormBase -w 400 -h 200 -text "ADD RECORD:"
            $usrlabel = New-Object System.Windows.Forms.Label
            $usrlabel.Location = New-Object System.Drawing.Size(10,30) 
            $usrlabel.Size = New-Object System.Drawing.Size(100,20) 
            $usrlabel.Text = "User:"
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

            Copy-Item -Path "$repopath\PatrolDB.csv.AES" -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
            $stringa = Get-Content "$repopath\crypto.key"
            $key = ConvertTo-SecureString $stringa -AsPlainText -Force
            Unprotect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Algorithm AES -Key $key -RemoveSource

            $string | Out-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv" -Encoding ASCII -Append

            net stop workstation /y > $null
            net start workstation > $null
            Start-Sleep 5
            $ErrorActionPreference = 'Stop'
            Try {
                New-PSDrive -Name P -PSProvider FileSystem -Root $repopath -Credential $logo > $null
            }
            Catch { 
                $errormsg = [System.Windows.MessageBox]::Show("Unable to access, check your credentials",'ERROR','Ok','Error')
            }
            $ErrorActionPreference = 'Inquire'
    
            $stringa = Get-Content 'P:\crypto.key'
            $key = ConvertTo-SecureString $stringa -AsPlainText -Force
            Protect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv" -Algorithm AES -Key $key -RemoveSource
            Remove-Item -Path "P:\PatrolDB.csv.AES"
            Copy-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Destination "P:\PatrolDB.csv.AES"
            Remove-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
    
            Remove-PSDrive -Name P

        }

        if ($check_user.Checked) {
            Copy-Item -Path "$repopath\PatrolDB.csv.AES" -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
            $stringa = Get-Content "$repopath\crypto.key"
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

            $formlist = FormBase -w 400 -h 200 -text "USERS LIST"
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
            $DropDownLabel.Text = "Select user"
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
            Write-Host -ForegroundColor Cyan "Script avalaible for $selected"
            $lista_risultato | sort
                
            Remove-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv"
        }

        if ($check_script.Checked) {
            Copy-Item -Path "$repopath\PatrolDB.csv.AES" -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
            $stringa = Get-Content "$repopath\crypto.key"
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

            $formlist = FormBase -w 400 -h 200 -text "SCRIPTS LIST"
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
            $DropDownLabel.Text = "Select script"
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
            Write-Host -ForegroundColor Cyan "Users granted to run $selected"
            $lista_risultato | sort

            Remove-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv"
        }

        if ($importa.Checked) {
            Write-Host "Import database..."

            net stop workstation /y > $null
            net start workstation > $null
            Start-Sleep 5
            $ErrorActionPreference = 'Stop'
            Try {
                New-PSDrive -Name P -PSProvider FileSystem -Root $repopath -Credential $logo > $null
            }
            Catch { 
                $errormsg = [System.Windows.MessageBox]::Show("Unable to access, check your credentials",'ERROR','Ok','Error')
            }
            $ErrorActionPreference = 'Inquire'
            
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
            [System.Windows.MessageBox]::Show("DB importato come $repopath\PatrolDB.csv.AES",'PATROL DB','Ok','Info') > $null
    
            Remove-PSDrive -Name P
        }
    
        if ($esporta.Checked) {
            Write-Host "Export database..."
            Copy-Item -Path "$repopath\PatrolDB.csv.AES" -Destination "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES"
            $stringa = Get-Content "$repopath\crypto.key"
            $key = ConvertTo-SecureString $stringa -AsPlainText -Force
            Unprotect-File "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv.AES" -Algorithm AES -Key $key -RemoveSource
    
            $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
            $SaveFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
            $SaveFileDialog.filter = "Comma separated value (*.csv)| *.csv"
            $SaveFileDialog.filename = $filename
            $SaveFileDialog.ShowDialog() | Out-Null
            $outfile = $SaveFileDialog.filename
            Move-Item -Path "C:\Users\$env:USERNAME\Desktop\PatrolDB.csv" -Destination $outfile
    
            [System.Windows.MessageBox]::Show("DB saved as $outfile",'PATROL DB','Ok','Info') > $null
        }
    }
} until ($ciclo -eq "Ok")
