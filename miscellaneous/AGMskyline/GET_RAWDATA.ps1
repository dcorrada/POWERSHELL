# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\miscellaneous\\AGMskyline\\GET_RAWDATA\.ps1$" > $null
$workdir = $matches[1]
<# alternative for testing
$workdir = Get-Location
$workdir = $workdir.Path
#>

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# tables dialog
$swlist = @()
$form_panel = FormBase -w 350 -h 700 -text "TABLES"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 20 -text "ADcomputers"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "ADusers"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 80 -text "AzureDevices" # migrare codice su Graph
$swlist += CheckBox -form $form_panel -checked $false -enabled $false -x 20 -y 110 -text "AzureDevices_nologin"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 140 -text "CheckinFrom"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 170 -text "DLmembers" # migrare codice su Graph
$swlist += CheckBox -form $form_panel -checked $false -enabled $false -x 20 -y 200 -text "DymoLabel" # tabella redatta manualmente
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 230 -text "EstrazioneAsset"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 260 -text "EstrazioneUtenti"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 290 -text "GFIparsed"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 320 -text "o365licenses" # migrare codice su Graph
$swlist += CheckBox -form $form_panel -checked $false -enabled $false -x 20 -y 350 -text "o365licenses_nologin"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 380 -text "PwdExpire"
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 410 -text "SchedeAssunzione"
$swlist += CheckBox -form $form_panel -checked $false -x 20 -y 440 -text "SchedeSIM" # chiedere a Max l'excel aggiornato
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 470 -text "SchedeTelefoni" 
$swlist += CheckBox -form $form_panel -checked $false -enabled $false -x 20 -y 500 -text "ThirdPartiesLicenses" # tabella redatta manualmente
$swlist += CheckBox -form $form_panel -checked $true -x 20 -y 530 -text "TrendMicroparsed"
$swlist += CheckBox -form $form_panel -checked $false -enabled $false -x 20 -y 560 -text "Xrefs" # tabelle create in fase di update
OKButton -form $form_panel -x 100 -y 600 -text "Ok"  | Out-Null
$form_panel.ShowDialog() | Out-Null

foreach ($item in $swlist) {
    if ($item.Checked) {
        $scriptfile = $workdir + '\miscellaneous\AGMskyline\' + $item.Text + '.ps1'
        Clear-Host
        Write-Host -ForegroundColor Yellow "Launching <$scriptfile>..."
        PowerShell.exe -file "$scriptfile"
    }
}

