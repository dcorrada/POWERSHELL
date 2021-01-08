# this is a test script to check account credentials

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
Import-Module -Name "$workdir\Modules\Forms.psm1"

 $form_PWD = New-Object System.Windows.Forms.Form
 $form_PWD.Text = "LOGIN"
 $form_PWD.Size = "400,250"
 $form_PWD.StartPosition = 'CenterScreen'
 $form_PWD.Topmost = $true
 $label = New-Object System.Windows.Forms.Label
 $label.Location = New-Object System.Drawing.Size(10,20) 
 $label.Size = New-Object System.Drawing.Size(300,20) 
 $label.Text = "Insert your credentials"
 $form_PWD.Controls.Add($label)
 $usrlabel = New-Object System.Windows.Forms.Label
 $usrlabel.Location = New-Object System.Drawing.Size(10,50) 
 $usrlabel.Size = New-Object System.Drawing.Size(100,20) 
 $usrlabel.Text = "Username:"
 $form_PWD.Controls.Add($usrlabel)
 $textBox = New-Object System.Windows.Forms.TextBox
 $textBox.Location = New-Object System.Drawing.Point(130,50)
 $textBox.Size = New-Object System.Drawing.Size(150,20)
 $form_PWD.Add_Shown({$textBox.Select()})
 $form_PWD.Controls.Add($textBox)
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
 $OKButton = New-Object System.Windows.Forms.Button
 $OKButton.Location = "100,120"
 $OKButton.Size = '100,30'
 $OKButton.Text = "Ok"
 $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
 $form_PWD.AcceptButton = $OKButton
 $form_PWD.Controls.Add($OKButton)
 $result = $form_PWD.ShowDialog()

$principalContext = [System.DirectoryServices.AccountManagement.PrincipalContext]::new([System.DirectoryServices.AccountManagement.ContextType]'Machine',$env:COMPUTERNAME)
$principalContext.ValidateCredentials($textBox.Text,$MaskedTextBox.Text)

Pause