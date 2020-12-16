<#
Name......: Grep-Replace.ps1
Version...: 20.2.1

Author....: Dario CORRADA

Questo script cerca ricorsivamente nei file .ps1 di una cartella una stringa specifica e la sostituisce

+++ UPDATES +++

[2019-10-08  CORRADA] 
Vedi GIT
#>

# header
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

# chiedo il percoso da cercare
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.RootFolder = "MyComputer"
$foldername.ShowDialog() > $null

# chiedo la chiave di ricerca e la stringa da sostituire
Import-Module -Name 'C:\Users\dario.corrada\OneDrive - AGM Solutions\POWERSHELL\Moduli_PowerShell\Forms.psm1'
$form_EXP = FormBase -w 350 -h 200 -text "TROVA E SOSTITUISCI"
$searchlabel = New-Object System.Windows.Forms.Label
$searchlabel.Location = New-Object System.Drawing.Point(10,20)
$searchlabel.Size = New-Object System.Drawing.Size(170,30)
$searchlabel.Text = "Cerca:"
$form_EXP.Controls.Add($searchlabel)
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(180,20)
$searchBox.Size = New-Object System.Drawing.Size(120,30)
$form_EXP.Controls.Add($searchBox)
$replacelabel = New-Object System.Windows.Forms.Label
$replacelabel.Location = New-Object System.Drawing.Point(10,50)
$replacelabel.Size = New-Object System.Drawing.Size(170,30)
$replacelabel.Text = "Sostituisci (vuoto Cerca solo):"
$form_EXP.Controls.Add($replacelabel)
$replaceBox = New-Object System.Windows.Forms.TextBox
$replaceBox.Location = New-Object System.Drawing.Point(180,50)
$replaceBox.Size = New-Object System.Drawing.Size(120,30)
$form_EXP.Controls.Add($replaceBox)
OKButton -form $form_EXP -x 75 -y 90 -text "Ok"
$form_EXP.Add_Shown({$replaceBox.Select()})
$form_EXP.Add_Shown({$searchBox.Select()})
$result = $form_EXP.ShowDialog()


$filelist = Get-ChildItem -Recurse -Path $foldername.SelectedPath -Filter "*.ps1"
foreach ($infile in $filelist) {
    Write-Host -ForegroundColor Cyan -NoNewline "Parsing" $infile.FullName
    if ($replaceBox.Text -eq "") {
        Write-Host " "
        $found = Get-content -path $infile.FullName | Select-String -pattern $searchBox.Text -encoding ASCII -CaseSensitive
        foreach ($newline in $found) {
            Write-Host $newline
        }
    } else {
        
        ((Get-Content -path $infile.FullName -Raw) -replace $searchBox.Text,$replaceBox.Text) | Set-Content -Path $infile.FullName
        Write-Host -ForegroundColor Green " DONE"
    }
}
Pause


