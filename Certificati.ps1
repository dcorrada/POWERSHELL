<#
Name......: Certificati.ps1
Version...: 20.12.1
Author....: Dario CORRADA

Questo script richiede o rimuove i certificati utente
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
Import-Module -Name "$workdir\Moduli_PowerShell\Forms.psm1"

# form di scelta 
$form_modalita = FormBase -w 300 -h 200 -text "CERTIFICATI"
$togli = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "Rimuovi certificato"
$metti  = RadioButton -form $form_modalita -checked $false -x 30 -y 60 -text "Importa certificato"
OKButton -form $form_modalita -x 90 -y 120 -text "Ok"
$result = $form_modalita.ShowDialog()

if ($result -eq "OK") {
    if ($togli.Checked) {
        # elenco i certificati installati
        $cert_list = Get-ChildItem cert:CurrentUser\My
        $opt_list = @()
        foreach ($cert in $cert_list) {
            $opt_list += $cert.Subject
        }
        $formlist = FormBase -w 400 -h 200 -text "CERTIFICATI"
        $DropDown = new-object System.Windows.Forms.ComboBox
        $DropDown.Location = new-object System.Drawing.Size(10,60)
        $DropDown.Size = new-object System.Drawing.Size(350,30)
        $font = New-Object System.Drawing.Font("Arial", 12)
        $DropDown.Font = $font
        foreach ($elem in ($opt_list | sort)) {
            $DropDown.Items.Add($elem)  > $null
        }
        $formlist.Controls.Add($DropDown)
        $DropDownLabel = new-object System.Windows.Forms.Label
        $DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
        $DropDownLabel.size = new-object System.Drawing.Size(500,30) 
        $DropDownLabel.Text = "Scegliere il certificato"
        $formlist.Controls.Add($DropDownLabel)
        OKButton -form $formlist -x 100 -y 100 -text "Ok"
        $formlist.Add_Shown({$DropDown.Select()})
        $result = $formlist.ShowDialog()
        $selected = $DropDown.Text

        # rimuovo il certificato
        foreach ($cert in $cert_list) { 
            if ($cert.Subject -match $selected) {
                Remove-Item -path $cert.PSPath -recurse -Force
            }
        }

    } elseif ($metti.Checked) {
        $answ = [System.Windows.MessageBox]::Show("Opzione non ancora disponibile",'TODO','Ok','Info')
<#
        $certreq_log = certreq -q -enroll -user [nomeutente]

        Questa riga importa automaticamente un certificato emesso da un utente.
        
        Leggersi la documentazione relativa al comando certreq su:
        https://docs.microsoft.com/it-it/windows-server/administration/windows-commands/certreq_1  
#>
    }
}