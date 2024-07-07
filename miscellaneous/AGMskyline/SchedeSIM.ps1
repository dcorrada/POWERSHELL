# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\miscellaneous\\AGMskyline\\SchedeSIM\.ps1$" > $null
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

# carico il file Excel graezzo gestito da Max
$answ = [System.Windows.MessageBox]::Show("Disponi di un file Excel aggiornato?",'INFILE','YesNo','Warning')
if ($answ -eq "No") {    
    Write-Host -ForegroundColor red "Aborting..."
    Start-Sleep -Seconds 3
    Exit
}
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Open File"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Excel file (*.xls)| *.xls'
$OpenFileDialog.ShowDialog() | Out-Null
$infile = $OpenFileDialog.filename

$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-SchedeSIM.csv'
#'ID;FULLNAME;PHONENUM;ACTIVATION;NOTES' | Out-File $outfile -Encoding utf8

$aexcel = New-Object -ComObject Excel.Application  
$abook = $aexcel.Workbooks.Open($infile)  
$asheet = $abook.Sheets.Item(1) 

$i = 2
$totrec = ($asheet.UsedRange.Rows).Count
$parsebar = ProgressBar
do {
    $phonenum = $asheet.Cells.Item($i,2).Text
    $phonenum = $phonenum -replace '^\s+', ''
    $phonenum = $phonenum -replace '\s+$', ''
    $fullname = $asheet.Cells.Item($i,3).Text + ' ' + $asheet.Cells.Item($i,4).Text
    $fullname = $fullname -replace '\(', ''
    $fullname = $fullname -replace 'TO\)', ''
    $fullname = $fullname -replace '\s\s\s*', ' '
    $attivazione = $asheet.Cells.Item($i,7).Text | Get-Date  -f yyyy-MM-dd
    $note = $asheet.Cells.Item($i,8).Text

    $string = ("AGM{0:d5};{1};{2};{3};{4}" -f (($i-1),$fullname,$phonenum,$attivazione,$note))
    $string = $string -replace ';\s*;', ';NULL;'
    $string = $string -replace ';+\s*$', ';NULL'
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
} while (!(($asheet.Cells.Item($i,1).Text) -eq '') -and ($i -lt ($asheet.UsedRange.Rows).Count))
$parsebar[0].Close()
$abook.close()
$aexcel.Quit()
