<#
Name......: Grep-Replace.ps1
Version...: 20.2.1
Author....: Dario CORRADA

This script recursively look for .ps1 files, grep and replace a string inside them
#>

# just pipe more than single "Split-Path" if the script maps to nested subfolders
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent

# header
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# asking search path
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") > $null
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.RootFolder = "MyComputer"
$foldername.ShowDialog() > $null

# define search and replace strings
$form_EXP = FormBase -w 350 -h 200 -text "GREP AND REPLACE"
$searchlabel = New-Object System.Windows.Forms.Label
$searchlabel.Location = New-Object System.Drawing.Point(10,20)
$searchlabel.Size = New-Object System.Drawing.Size(170,30)
$searchlabel.Text = "Search:"
$form_EXP.Controls.Add($searchlabel)
$searchBox = New-Object System.Windows.Forms.TextBox
$searchBox.Location = New-Object System.Drawing.Point(180,20)
$searchBox.Size = New-Object System.Drawing.Size(120,30)
$form_EXP.Controls.Add($searchBox)
$replacelabel = New-Object System.Windows.Forms.Label
$replacelabel.Location = New-Object System.Drawing.Point(10,50)
$replacelabel.Size = New-Object System.Drawing.Size(170,30)
$replacelabel.Text = "Replace (empty = only Search):"
$form_EXP.Controls.Add($replacelabel)
$replaceBox = New-Object System.Windows.Forms.TextBox
$replaceBox.Location = New-Object System.Drawing.Point(180,50)
$replaceBox.Size = New-Object System.Drawing.Size(120,30)
$form_EXP.Controls.Add($replaceBox)
OKButton -form $form_EXP -x 75 -y 90 -text "Ok" | Out-Null
$form_EXP.Add_Shown({$replaceBox.Select()})
$form_EXP.Add_Shown({$searchBox.Select()})
$result = $form_EXP.ShowDialog()


$filelist = Get-ChildItem -Recurse -Path $foldername.SelectedPath -Filter "*.ps1"
foreach ($infile in $filelist) {
    if ($replaceBox.Text -eq "") {
        $found = Get-content -path $infile.FullName | Select-String -pattern $searchBox.Text -encoding ASCII -CaseSensitive
        if ($found.Count -ge 1) {
            Write-Host -ForegroundColor Cyan "$($infile.FullName)"
            foreach ($newline in $found) {
                Write-Host "  $newline"
            }
        }
    } else {
        Write-Host -ForegroundColor Cyan -NoNewline "Parsing" $infile.FullName
        ((Get-Content -path $infile.FullName -Raw) -replace $searchBox.Text,$replaceBox.Text) | Set-Content -Path $infile.FullName
        Write-Host -ForegroundColor Green " DONE"
    }
}
Pause
