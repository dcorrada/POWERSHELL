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

# elevated script execution with admin privileges
$ErrorActionPreference= 'Stop'
try {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if ($testadmin -eq $false) {
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
        exit $LASTEXITCODE
    }
}
catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}
$ErrorActionPreference= 'Inquire'

# just pipe more than single "Split-Path" if the script maps to nested subfolders
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent | Split-Path -Parent

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# carico il csv estratto dal tenant web
$answ = [System.Windows.MessageBox]::Show("Disponi di un file CSV con separatore <,> ?",'INFILE','YesNo','Warning')
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
# sostituisco l'header del file per non avere lettere accentate e spazi
$Heather = 'displayName', 'accountEnabled', 'operatingSystem', 'OSVersion', 'joinType',
    'registeredOwners', 'userNames', 'mdmDisplayName', 'isCompliant', 'registrationTime',
    'LastSignIn', 'deviceId', 'isManaged', 'objectId', 'profileType', 'systemLabels', 'model'
$A = Get-Content -Path $infile
$A = $A[1..($A.Count - 1)]
$A | Out-File -FilePath $infile
$rawdata = Import-Csv -Path $infile -Header $Heather

$parseddata = @{}
$tot = $rawdata.Count
$i = 0
$parsebar = ProgressBar
foreach ($item in $rawdata) {
    $i++
    $akey = $item.displayName
    if ($parseddata.ContainsKey($akey)) {
        $uptime = $item.LastSignIn | Get-Date -f yyy-MM-dd
        if ($uptime -gt $parseddata[$akey].LastLogon) {
            $parseddata[$akey].LastLogon = $uptime
            $parseddata[$akey].OSType = $item.operatingSystem
            $parseddata[$akey].OSVersion = $item.OSVersion
            $parseddata[$akey].Owner = $item.registeredOwners
            $parseddata[$akey].Email = $item.userNames
        }
    } else {
        $parseddata[$akey] = @{
            'LastLogon' = $item.LastSignIn | Get-Date -f yyy-MM-dd
            'OSType' = $item.operatingSystem
            'OSVersion' = $item.OSVersion
            'HostName' = $item.displayName
            'Owner' = $item.registeredOwners
            'Email' = $item.userNames 
        }
    }

    # progress
    $percent = ($i / $tot)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Record {0} out of {1} parsed [{2}%]" -f ($i, $tot, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}
$parsebar[0].Close()


# writing output file
Write-Host -NoNewline "Writing output file... "
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-AzureDevices.csv'


$i = 1
$totrec = $parseddata.Count
$parsebar = ProgressBar
foreach ($item in $parseddata.Keys) {
    $string = ("AGM{0:d5};{1};{2};{3};{4};{5};{6}" -f ($i,$parseddata[$item].LastLogon,$parseddata[$item].OSType,$parseddata[$item].OSVersion,$parseddata[$item].HostName,$parseddata[$item].Owner,$parseddata[$item].Email))
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
}
$parsebar[0].Close()
Write-Host -ForegroundColor Green "DONE"

