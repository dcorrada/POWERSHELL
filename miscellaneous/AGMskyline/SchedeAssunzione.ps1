# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\miscellaneous\\AGMskyline\\SchedeAssunzione\.ps1$" > $null
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

# carico il modulo per importare file excel
$ErrorActionPreference= 'Stop'
try {
    Import-Module ImportExcel
} catch {
    Install-Module ImportExcel -Confirm:$False -Force
    Import-Module ImportExcel
}
$ErrorActionPreference= 'Inquire'

# percorsi di rete
$target_paths = @{
    'ASSUNTO' = '\\192.168.2.251\Share\AREA_SCAMBIO\HR\ASSUNZIONI'
    'CESSATO' = '\\192.168.2.251\Share\AREA_SCAMBIO\HR\CESSAZIONI'
}

# writing output file
Write-Host "Writing output file... "
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-SchedeAssunzione.csv'

$i = 1
foreach ($status in $target_paths.Keys) {
    # recupero la lista delle schede
    $filelist = Get-ChildItem -Path ($target_paths[$status] + '\*.xlsx') -Name

    Write-Host -ForegroundColor Blue "*** ANALISI SCHEDE $status ***"
    Start-Sleep 2

    $tot = $filelist.Count
    $bari = 0
    $parsebar = ProgressBar
    foreach ($item in $filelist) {
        $afile = $target_paths[$status] + '\' + $item
        # Write-Host "Processing [$afile]... "
        $rawdata = Import-Excel -Path $afile -NoHeader
        $nextstart = $false
        foreach ($slot in $rawdata) {
            if ($nextstart -eq $true) {
                $ErrorActionPreference= 'Stop'
                try {
                    $started = $slot.P1  | Get-Date -format "yyyy-MM-dd"
                } catch {
                    $started = 'NULL'
                }
                $ErrorActionPreference= 'Inquire'
                $nextstart = $false
            } else {
                if ($slot.P1 -eq 'NOME') {
                    $nome = $slot.P2
                } elseif ($slot.P1 -eq 'COGNOME') {
                    $cognome = $slot.P2
                } elseif ($slot.P1 -eq 'EMAIL AZIENDALE') {
                    $email = $slot.P2
                } elseif ($slot.P1 -eq 'AREA DI LAVORO') {
                    $role = $slot.P2
                } elseif ($slot.P1 -eq 'UTENTE INTERNO / ESTERNO') {
                    $scope = $slot.P2
                } elseif ($slot.P1 -eq 'SEDE AGM DI RIFERIMENTO') {
                    $location = $slot.P2
                } elseif ($slot.P1 -eq 'DATA INIZIO') {
                    $nextstart = $true
                }  elseif (($slot.P1 -eq 'UTENTE') -and !($slot.P2 -match '@agmsolutions.net')) {
                    $usernamed = $slot.P2
                }
            }            
        }
        $TextInfo = (Get-Culture).TextInfo
        $fullname = $TextInfo.ToTitleCase("{0} {1}" -f ($nome.ToLower(),$cognome.ToLower()))

        $string = ("AGM{0:d5};{1};{2};{3};{4};{5};{6};{7};{8}" -f ($i,$fullname,$email,$usernamed,$role,$scope,$status,$location.ToUpper(),$started))
        $string = $string -replace ';\s*;', ';NULL;'
        $string = $string -replace ';+\s*$', ';NULL'
        $string = $string -replace ';', '";"'
        $string = '"' + $string + '"'
        $string = $string -replace '"NULL"', 'NULL'
        $string | Out-File $outfile -Encoding utf8 -Append
        $i++

        # progress
        $bari++
        $percent = ($bari / $tot)*100
        if ($percent -gt 100) {
            $percent = 100
        }
        $formattato = '{0:0.0}' -f $percent
        [int32]$progress = $percent   
        $parsebar[2].Text = ("Record {0} out of {1} parsed [{2}%]" -f ($bari, $tot, $formattato))
        if ($progress -ge 100) {
            $parsebar[1].Value = 100
        } else {
            $parsebar[1].Value = $progress
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    $parsebar[0].Close()
}