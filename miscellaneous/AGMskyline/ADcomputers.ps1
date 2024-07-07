# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\ADcomputers\.ps1$" > $null
$workdir = $matches[1]
<# alternative for testing
$workdir = Get-Location
$workdir = $workdir.Path
#>

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Forms.psm1"

# import Active Directory module
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory }

# collecting data
Write-Host -NoNewline "Retrieving computer list..."
$computer_list = Get-ADComputer -Filter * -Property *
Write-Host -ForegroundColor Green 'DONE'
$fetched_data = @()
$i = 1
$totrec = $computer_list.Count
$parsebar = ProgressBar

foreach ($computer_name in $computer_list.Name) {

    $matches = ('', 'NULL')
    $infopc = Get-ADComputer -Identity $computer_name -Properties *
    $infopc.CanonicalName -match "/(.+)/$computer_name$" > $null
    $ou = $matches[1]

    # empty dates
    $ErrorActionPreference= 'Stop'
    try {
        $logondate = $infopc.LastLogonDate | Get-Date -format "yyyy-MM-dd"
    } catch {
        $logondate = 'NULL'
    }
    $ErrorActionPreference= 'Inquire'

    $fetched_record = @{
        NAME = $computer_name
        OS = $infopc.OperatingSystem + ' [' + $infopc.OperatingSystemVersion + ']'
        UPDATED = $logondate
        OU = $ou
    }

    $fetched_data += $fetched_record
    $i++

    # progress
    $percent = (($i-1) / $totrec)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Collecting {0} out of {1} records [{2}%]" -f (($i-1), $totrec, $formattato))
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
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-ADcomputers.csv'


$i = 1
$totrec = $fetched_data.Count
$parsebar = ProgressBar
foreach ($item in $fetched_data) {
    $string = ("AGM{0:d5};{1};{2};{3};{4}" -f ($i,$item.NAME,$item.OU,$item.OS,$item.UPDATED))
    $string = $string -replace ';\s*;', ';NULL;'
    $string = $string -replace ';+\s*$', ';NULL'
    $string = $string -replace ';"\s\[\]";', ';NULL;'
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
