<#
Name......: WiFi_Profiles.ps1
Version...: 20.12.1
Author....: Dario CORRADA

This script import or export wireless profiles
#>

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$workdir = Get-Location
$workdir -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Locales$" > $null
$repopath = $matches[1]
Import-Module -Name "$repopath\Modules\Forms.psm1"

# dialog box 
$form_modalita = FormBase -w 370 -h 220 -text "WIFI PROFILES"
$esporta = RadioButton -form $form_modalita -checked $true -x 30 -y 20 -text "Export profiles"
$importa  = RadioButton -form $form_modalita -checked $false -x 30 -y 60 -text "Import profiles"
OKButton -form $form_modalita -x 90 -y 120 -text "Ok" | Out-Null
$result = $form_modalita.ShowDialog()

if ($result -eq "OK") {
    if ($esporta.Checked) {
        # export profiles
        $dest = "C:\Users\$env:USERNAME\Desktop\wifi_profiles"
        Write-Host "Export WiFi profiles..."
        New-Item -ItemType directory -Path $dest
        netsh wlan export profile key=clear folder=$dest
        [System.Windows.MessageBox]::Show("Profiles exported to $dest",'EXPORT','Ok','Info')
    } elseif ($importa.Checked) {
        # import profiles
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            RootFolder            = "MyComputer"
            Description           = "Select profiles folder"
        }
        $Null = $FolderBrowser.ShowDialog()
        $source = $FolderBrowser.SelectedPath
        Write-Host "Import WiFi profiles..."
        $wifi_profile_list = Get-ChildItem $source
        $exclude_list = @('foo.xml', 'bar.xml', 'baz.xml')
        foreach ($wifi_profile in $wifi_profile_list) {
            if ($exclude_list -contains $wifi_profile) {
                Write-Host "Skip $wifi_profile"
            } else {
                netsh wlan add profile filename="$source\$wifi_profile" user=current
            }
        }
        [System.Windows.MessageBox]::Show("Profiles imported from $source",'IMPORT','Ok','Info')
    }
}
