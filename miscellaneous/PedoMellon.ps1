param ([string]$UsrStr)

<#
Name......: PedoMellon.ps1
Version...: 24.09.a
Author....: Dario CORRADA

This script is a password generator: the strings generated should be easy to keep 
in mind but complex enough to decipher, as well
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# check execution policy
foreach ($item in (Get-ExecutionPolicy -List)) {
    if(($item.Scope -eq 'LocalMachine') -and ($item.ExecutionPolicy -cne 'Bypass')) {
        Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:`n"
        Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force`n"
        Write-Host -NoNewline "Afterwards restart this script. "
        Pause
        Exit
    }
}

# just pipe more than single "Split-Path" if the script maps to nested subfolders
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent

# graphical stuff
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

<# *******************************************************************************
                                    DIALOG
******************************************************************************* #>
if ([string]::IsNullOrEmpty($UsrStr)) {
    $continueBrowsing = $true
    while ($continueBrowsing) {
        $TheDialog = FormBase -w 850 -h 230 -text "PEDO MELLON A MINNO"

        # string area
        $UsrBox = TxtBox -form $TheDialog -x 20 -y 20 -h 35 -w 200 -text "Insert here a username..."
        $UsrBox.Font = [System.Drawing.Font]::new("Arial", 11)
        $PwdBox = TxtBox -form $TheDialog -x 20 -y 60 -h 35 -w 200 -text "here there is a pswd!" # sostituire con la variabile contenente la passsword
        $PwdBox.ReadOnly = $true
        $PwdBox.Font = [System.Drawing.Font]::new("Courier New", 11)

        # methods area
        <# lunghezza, tipologia caratteri, metodi, ... #>

        # buttons
        RETRYButton -form $TheDialog -x 20 -y 100 -text "Generate" | Out-Null
        $CopyBut = RETRYButton -form $TheDialog -x 120 -y 100 -text "Copy"
        $CopyBut.DialogResult = [System.Windows.Forms.DialogResult]::IGNORE
        OKButton -form $TheDialog -x 20 -y 140 -w 200 -text "Quit" | Out-Null

        $ButtHead = $TheDialog.ShowDialog()

        if ($ButtHead -eq 'IGNORE') { 
            [System.Windows.Forms.Clipboard]::SetText($PwdBox.Text)
        } elseif ($ButtHead -eq 'RETRY') {
            # generate another password
        } elseif ($ButtHead -eq 'OK') {
            $continueBrowsing = $false
        }
    }

<# *******************************************************************************
                                 HEADLESS
******************************************************************************* #>
} else {
    <# generatore senza dialog grafica, per integrare lo script in altri #>
}
