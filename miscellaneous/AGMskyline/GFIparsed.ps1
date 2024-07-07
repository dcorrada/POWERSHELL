# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\miscellaneous\\AGMskyline\\GFIparsed\.ps1$" > $null
$workdir = $matches[1]
<# alternative for testing
$workdir = Get-Location
$workdir = $workdir.Path
#>

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# carico il csv preprocessato (caricare l'estrazione dal filtro di GFI in excel ed esportare il file nel formato "CSV (delimitato da separatore di elenco)")
$answ = [System.Windows.MessageBox]::Show("Disponi di un file CSV?",'INFILE','YesNo','Warning')
if ($answ -eq "No") {    
    Write-Host -ForegroundColor red "Aborting..."
    Start-Sleep -Seconds 3
    Exit
}
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Open File"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Downloads"
$OpenFileDialog.filter = 'CSV file (*.csv)| *.csv'
$OpenFileDialog.ShowDialog() | Out-Null
$infile = $OpenFileDialog.filename

$rawdata = Import-Csv -Path $infile
$chomp = @()
$selected_fields = ('OS', 'Descrizione', 'Versione agente', 'Nome utente', 'Dispositivo', 'Antivirus gestito (Bitdefender)')
foreach ($item in $rawdata) {
    $arecord = @{}
    foreach ($akey in $selected_fields) {
        $arecord["$akey"] = $item."$akey"
    }
    $chomp += $arecord
}

# writing output file
Write-Host -NoNewline "Writing output file... "
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-GFIparsed.csv'


$i = 1
$totrec = $chomp.Count
$parsebar = ProgressBar
foreach ($item in $chomp) {
    $matches = @()
    $item['Nome utente'] -match "^([a-zA-Z_\-\.\\\s0-9:]+)\\([a-zA-Z\.0-9]+)$" > $null
    if ($matches[2]) {
        $usernamed = $matches[2]
    } else {
        $usernamed = 'NULL'
    }
    $string = ("AGM{0:d5};{1};{2};{3};{4};{5};{6}" -f ($i,$usernamed,$item['Dispositivo'],$item['OS'],$item['Descrizione'],$item['Versione agente'],$item['Antivirus gestito (Bitdefender)']))
    $string = $string -replace ';\s*;', ';NULL;'
    $string = $string -replace ';+\s*$', ';NULL'
    $string = $string -replace ';\s\[\];', ';NULL;'
    $string = $string -replace ';', '";"'
    $string = '"' + $string + '"'
    $string = $string -replace '"NULL"', 'NULL'
    $string | Out-File $outfile -Encoding utf8 -Append
    $i++

    # progress
    $percent = (($i-1) / $totrec)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Writing {0} out of {1} records [{2}%]" -f (($i-1), $totrec, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()    
}
$parsebar[0].Close()
Write-Host -ForegroundColor Green "DONE"
