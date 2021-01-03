# functions for dialog boxes


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