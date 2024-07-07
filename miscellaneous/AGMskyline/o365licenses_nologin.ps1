# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\miscellaneous\\AGMskyline\\o365licenses_nologin\.ps1$" > $null
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

# carico il csv estratto dal tenant web
$answ = [System.Windows.MessageBox]::Show("Disponi di un file CSV?",'INFILE','YesNo','Warning')
if ($answ -eq "No") {    
    Write-Host -ForegroundColor red "Aborting..."
    Start-Sleep -Seconds 1
    Exit
}
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Open File"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Downloads"
$OpenFileDialog.filter = 'CSV file (*.csv)| *.csv'
$OpenFileDialog.ShowDialog() | Out-Null
$infile = $OpenFileDialog.filename
# sostituisco l'header del file per non avere lettere accentate e spazi
$Heather = 'blocked', 'city', 'state', 'dept', 'DirSyncEnabled', 'DisplayName',
    'Fax', 'name', 'LastDirSync', 'surname', 'LastPasswordSet', 'details', 'licenses',
    'phone', 'MetadataTokenOATH', 'ObjID', 'Office', 'NeverExpire', 'phone2', 'CAP',
    'preferred', 'language','proxy', 'deploy', 'eliminazione_temporanea', 'country',
    'addres', 'pswdComplex', 'title', 'zone', 'UPN', 'created'
$A = Get-Content -Path $infile
$A = $A[1..($A.Count - 1)]
$A | Out-File -FilePath $infile
$UsrRawdata = Import-Csv -Path $infile -Header $Heather

# initialize dataframe for collecting data
$parseddata = @{}

$tot = $UsrRawdata.Count
$usrcount = 0
$parsebar = ProgressBar
Write-Host "Collecting Data..."
foreach ($User in $UsrRawdata) {
    $usrcount ++

    if ([string]::IsNullOrEmpty($User.licenses)) {
        Write-Host -ForegroundColor Yellow "[SKIP] Nessuna licenza assegnata a <$($User.UPN)>"
    } else {
        $parseddata[$User.UPN] = @{
            'fullname' = $User.DisplayName
            'email' = $User.UPN
            'licenza' = $User.licenses.Split('+')
            'start' = ($User.created | Get-Date -format "yyyy/MM/dd")
        }
    }

    # progress
    $percent = ($usrcount / $tot)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Record {0} out of {1} parsed [{2}%]" -f ($usrcount, $tot, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}
Write-Host -ForegroundColor Green " DONE"
$parsebar[0].Close()


# writing output file
Write-Host -NoNewline "Writing output file... "
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-o365licenses.csv'


$i = 1
$totrec = $parseddata.Keys.Count
$parsebar = ProgressBar
foreach ($item in $parseddata.Keys) {
    foreach ($unalicenza in $parseddata[$item].licenza) {
        $string = ("AGM{0:d5};{1};{2};{3};{4}" -f ($i,$parseddata[$item].fullname,$parseddata[$item].email,$parseddata[$item].start,$unalicenza))
        $string = $string -replace ';\s*;', ';NULL;'
        $string = $string -replace ';+\s*$', ';NULL'
        $string = $string -replace ';', '";"'
        $string = '"' + $string + '"'
        $string = $string -replace '"NULL"', 'NULL'
        $string | Out-File $outfile -Encoding utf8 -Append
        $i++
    }

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
