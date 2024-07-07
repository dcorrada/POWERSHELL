# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\ADusers\.ps1$" > $null
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

Write-Host -NoNewline "Retrieving users list..."
$user_list = Get-ADUser -Filter * -Property *
Write-Host -ForegroundColor Green 'DONE'
$rawdata = @{}
$i = 1
$totrec = $user_list.Count
$parsebar = ProgressBar

foreach ($auser in $user_list) {
    $fullname = $auser.Name
    $usrname = $auser.SamAccountName
    $UPN = $auser.UserPrincipalName 
    $matches = ('', 'NULL')
    $auser.CanonicalName -match "(.+)/[a-zA-Z_\-\.\\\s0-9:]+$" > $null
    $ou = $matches[1]
    $ErrorActionPreference= 'Stop'
    try {
        $logondate = $auser.LastLogonDate | Get-Date -format "yyyy-MM-dd"
        $created = $auser.Created | Get-Date -format "yyyy-MM-dd"
    } catch {
        $logondate = 'NULL'
        $created = 'NULL'
    }
    $ErrorActionPreference= 'Inquire'
    if ($auser.Description -eq '') {
        $description = 'NULL'
    } else {
        $description = $auser.Description
    }
    
    $rawdata.$usrname = @{
        USRNAME = $usrname
        UPN = $UPN
        OU = $ou
        CREATED =  $created
        FULLNAME = $fullname
        LASTLOGON = $logondate
        DESCRIPTION = $description
    }
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
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-ADusers.csv'

$i = 1
$totrec = $rawdata.Count
$parsebar = ProgressBar
foreach ($usr in $rawdata.Keys) {
    $new_record = @(
        #$usr,
        $rawdata.$usr.USRNAME,
        $rawdata.$usr.UPN,
        $rawdata.$usr.FULLNAME,
        $rawdata.$usr.OU,
        $rawdata.$usr.CREATED,
        $rawdata.$usr.LASTLOGON,
        $rawdata.$usr.DESCRIPTION
    )
    $packed = [system.String]::Join(";", $new_record)

    $string = ("AGM{0:d5};{1}" -f ($i,$packed))
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
