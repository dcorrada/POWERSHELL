# funzioni per la gestione delle finestre


function FormBase {
    param ($w, $h, $text)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $text
    $form.Size = "$w,$h"
    $form.StartPosition = 'CenterScreen'
    $form.Topmost = $true

    return $form
}
Export-ModuleMember -Function FormBase

function RadioButton {
    param ($form, $checked, $x, $y, $text)

    $obj = New-Object System.Windows.Forms.RadioButton
    $obj.Location = "$x,$y"
    $obj.Size = '300,30'
    $obj.Checked = $checked
    $obj.Text = $text
    $form.Controls.Add($obj)

    return $obj
}
Export-ModuleMember -Function RadioButton

function CheckBox {
    param ($form, $checked, $x, $y, $text)

    $obj = New-Object System.Windows.Forms.CheckBox
    $obj.Location = "$x,$y"
    $obj.Size = '350,30'
    $obj.Checked = $checked
    $obj.Text = $text
    $form.Controls.Add($obj)

    return $obj
}
Export-ModuleMember -Function CheckBox

function OKButton {
    param ($form, $x, $y, $text)

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = "$x,$y"
    $OKButton.Size = '100,30'
    $OKButton.Text = $text
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)
}
Export-ModuleMember -Function OKButton

function LoginWindow {
    $form_PWD = New-Object System.Windows.Forms.Form
    $form_PWD.Text = "LOGIN"
    $form_PWD.Size = "400,250"
    $form_PWD.StartPosition = 'CenterScreen'
    $form_PWD.Topmost = $true

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
    $usr = $textBox.Text
    $pwd = ConvertTo-SecureString $MaskedTextBox.Text -AsPlainText -Force
    $login = New-Object System.Management.Automation.PSCredential($usr, $pwd)

    return $login
}
Export-ModuleMember -Function LoginWindow