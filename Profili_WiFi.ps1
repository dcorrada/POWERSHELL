<#
Name......: Profili_WiFi.ps1
Version...: 20.12.1
Author....: Dario CORRADA

Questo script importa o esporta profili wireless
#>

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
Import-Module -Name "$workdir\Moduli_PowerShell\Forms.psm1"

# scelta 
$form_modalita = FormBase -w 370 -h 220 -text "PROFILI WIFI"
$esporta = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "Esporta profili"
$importa  = RadioButton -form $form_modalita -checked $false -x 30 -y 60 -text "Importa profili"
OKButton -form $form_modalita -x 90 -y 120 -text "Ok"
$result = $form_modalita.ShowDialog()

if ($result -eq "OK") {
    if ($esporta.Checked) {
        # esporto i profili
        $dest = "C:\Users\$env:USERNAME\Desktop\wifi_profiles"
        Write-Host "Esporto profili Wifi..."
        New-Item -ItemType directory -Path $dest
        netsh wlan export profile key=clear folder=$dest
        [System.Windows.MessageBox]::Show("Profili esportati su $dest",'ESPORTA','Ok','Info')
    } elseif ($importa.Checked) {
        # importo i profili
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            RootFolder            = "MyComputer"
            Description           = "Seleziona cartella profili"
        }
        $Null = $FolderBrowser.ShowDialog()
        $source = $FolderBrowser.SelectedPath
        Write-Host "Importo profili Wifi..."
        $wifi_profile_list = Get-ChildItem $source
        $exclude_list = ("Wi-Fi-AGM.xml")
        foreach ($wifi_profile in $wifi_profile_list) {
            if ($exclude_list -contains $wifi_profile) {
                Write-Host "Skip $wifi_profile"
            } else {
                netsh wlan add profile filename="$source\$wifi_profile" user=current
            }
        }
        [System.Windows.MessageBox]::Show("Profili importati da $source",'ESPORTA','Ok','Info')
    }
}
