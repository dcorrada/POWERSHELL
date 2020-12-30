<#
Name......: COMPILA.ps1
Version...: 19.05.1
Author....:  CORRADA

Questo script e' un'interfaccia grafica per il compilatore
#>

$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms')
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'PowerShell script (*.ps1)| *.ps1'
$OpenFileDialog.ShowDialog() | Out-Null
$ps1file = $OpenFileDialog.filename

$ps1file -match "\\([a-zA-Z_\-\.\s0-9]+)\.ps1$" > $null
$filename = $matches[1]

$form = New-Object System.Windows.Forms.Form
$form.Text = "OPZIONI"
$form.Size = New-Object System.Drawing.Size(520,380)
$form.StartPosition = 'CenterScreen'
$font = New-Object System.Drawing.Font("Arial", 12)
$form.Font = $font

$label1 = New-Object System.Windows.Forms.Label
$label1.Location = New-Object System.Drawing.Point(10,20)
$label1.Size = New-Object System.Drawing.Size(500,30)
$label1.Text = "Company:"
$form.Controls.Add($label1)

$textBox1 = New-Object System.Windows.Forms.TextBox
$textBox1.Location = New-Object System.Drawing.Point(10,60)
$textBox1.Size = New-Object System.Drawing.Size(450,30)
$textBox1.Text = "AGM Solutions"
$form.Controls.Add($textBox1)
            
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,100)
$label2.Size = New-Object System.Drawing.Size(500,30)
$label2.Text = "Copyright:"
$form.Controls.Add($label2)

$textBox2 = New-Object System.Windows.Forms.TextBox
$textBox2.Location = New-Object System.Drawing.Point(10,140)
$textBox2.Size = New-Object System.Drawing.Size(450,30)
$textBox2.Text = "2020 -  CORRADA"
$form.Controls.Add($textBox2)

$label3 = New-Object System.Windows.Forms.Label
$label3.Location = New-Object System.Drawing.Point(10,180)
$label3.Size = New-Object System.Drawing.Size(500,30)
$label3.Text = "Version:"
$form.Controls.Add($label3)

$textBox3 = New-Object System.Windows.Forms.TextBox
$textBox3.Location = New-Object System.Drawing.Point(10,220)
$textBox3.Size = New-Object System.Drawing.Size(450,30)
$textBox3.Text = "20."
$form.Controls.Add($textBox3)

$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(100,280)
$OKButton.Size = New-Object System.Drawing.Size(75,30)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)

$form.Topmost = $true
$result = $form.ShowDialog()

$company = $textBox1.Text
$copyright = $textBox2.Text
$version = $textBox3.Text

[System.Reflection.Assembly]::LoadWithPartialName(�System.windows.forms�)
$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = �Executable (*.exe)| *.exe�
$OpenFileDialog.filename = $filename
$OpenFileDialog.ShowDialog() | Out-Null
$exefile = $OpenFileDialog.filename

$cmd = "C:\Users\$env:USERNAME\OneDrive - AGM Solutions\POWERSHELL\PS2EXE\ps2exe.ps1"
& $cmd -inputFile $ps1file -outputFile $exefile -verbose -company $company -copyright $copyright -version $version -noConfigfile



