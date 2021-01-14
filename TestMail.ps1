# script per mandare mail

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# import modules
$workdir = Get-Location
Import-Module -Name "$workdir\Modules\Forms.psm1"

$answ = [System.Windows.MessageBox]::Show("Configure for sending mail alerts?",'ALERTS','YesNo','Info')
if ($answ -eq "Yes") {    
    # dialog box
    $formail = New-Object System.Windows.Forms.Form
    $formail.Text = "CONFIG"
    $formail.Size = "500,270"
    $formail.StartPosition = 'CenterScreen'
    $formail.Topmost = $true
    $address = New-Object System.Windows.Forms.Label
    $address.Location = New-Object System.Drawing.Size(10,20) 
    $address.Size = New-Object System.Drawing.Size(120,20) 
    $address.Text = "Mail address:"
    $formail.Controls.Add($address)
    $addressbox = New-Object System.Windows.Forms.TextBox
    $addressbox.Location = New-Object System.Drawing.Point(130,20)
    $addressbox.Size = New-Object System.Drawing.Size(300,20)
    $formail.Add_Shown({$addressbox.Select()})
    $formail.Controls.Add($addressbox)
    $passwd = New-Object System.Windows.Forms.Label
    $passwd.Location = New-Object System.Drawing.Size(10,50) 
    $passwd.Size = New-Object System.Drawing.Size(120,20) 
    $passwd.Text = "Password:"
    $formail.Controls.Add($passwd)
    $passwdbox = New-Object System.Windows.Forms.MaskedTextBox
    $passwdbox.PasswordChar = '*'
    $passwdbox.Location = New-Object System.Drawing.Point(130,50)
    $passwdbox.Size = New-Object System.Drawing.Size(300,20)
    $formail.Add_Shown({$passwdbox.Select()})
    $formail.Controls.Add($passwdbox)
    $smtp = New-Object System.Windows.Forms.Label
    $smtp.Location = New-Object System.Drawing.Size(10,80) 
    $smtp.Size = New-Object System.Drawing.Size(120,20) 
    $smtp.Text = "SMTP server:"
    $formail.Controls.Add($smtp)
    $smtpbox = New-Object System.Windows.Forms.TextBox
    $smtpbox.Location = New-Object System.Drawing.Point(130,80)
    $smtpbox.Size = New-Object System.Drawing.Size(300,20)
    $formail.Add_Shown({$smtpbox.Select()})
    $formail.Controls.Add($smtpbox)
    $port = New-Object System.Windows.Forms.Label
    $port.Location = New-Object System.Drawing.Size(10,110) 
    $port.Size = New-Object System.Drawing.Size(120,20) 
    $port.Text = "Port:"
    $formail.Controls.Add($port)
    $portbox = New-Object System.Windows.Forms.TextBox
    $portbox.Location = New-Object System.Drawing.Point(130,110)
    $portbox.Size = New-Object System.Drawing.Size(300,20)
    $portbox.Text = '587'
    $formail.Add_Shown({$portbox.Select()})
    $formail.Controls.Add($portbox)
    $sslbox = New-Object System.Windows.Forms.CheckBox
    $sslbox.Location = New-Object System.Drawing.Point(130,140)
    $sslbox.Size = New-Object System.Drawing.Size(300,20)
    $sslbox.Text = "TLS/SSL authentication"
    $sslbox.Checked = $true
    $formail.Add_Shown({$sslbox.Select()})
    $formail.Controls.Add($sslbox)
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = "150,180"
    $OKButton.Size = '100,30'
    $OKButton.Text = "Ok"
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $formail.AcceptButton = $OKButton
    $formail.Controls.Add($OKButton)
    $result = $formail.ShowDialog()

    # setting credentials
    $usr = $addressbox.Text
    $pwd = ConvertTo-SecureString $passwdbox.Text -AsPlainText -Force
    $credential = New-Object System.Management.Automation.PSCredential($usr, $pwd)

    # define email content
    $subject = 'TestMail.ps1'
    $body = "Questa mail è stata mandata da uno script PowerShell"

    # sending email
    $ErrorActionPreference= 'Stop'
    Try {
        $SMTPClient = New-Object Net.Mail.SmtpClient($smtpbox.Text, $portbox.Text)
        if ($sslbox.Checked) {
            $SMTPClient.EnableSsl = $true
        }
        $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($addressbox.Text, $passwdbox.Text);
        $SMTPClient.Send($addressbox.Text, $addressbox.Text, $subject, $body)


        $ErrorActionPreference= 'Inquire'
    }
    Catch {
        Write-Output "`nError: $($error[0].ToString())"
        $answ = [System.Windows.MessageBox]::Show("Sending alert email failed",'WARNING','Ok','Warning')
    }   
}
