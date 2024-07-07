# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\EstrazioneUtenti\.ps1$" > $null
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

# accessing info
$uri_prefix = 'http://192.168.2.184/' # IP Snipe webserver
$token_file = $env:LOCALAPPDATA + '\SnipeIT.token'
$token_string = Get-Content $token_file

# header for request
$headers = @{
    'Authorization' = "Bearer $token_string"
    'Accept' = 'application/json'
    'Content-Type' = 'application/json'
}

# query parameters
$query_params = @(
    'limit=1000'
    'offset=0'
)
$uri_suffix = '?' + ($query_params -join '&')
$uri = $uri_prefix + 'api/v1/users' + $uri_suffix
$ErrorActionPreference= 'Stop'
Try {
    $rawdata = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("$($error[0].ToString())",'ABORTING','Ok','Error')
    exit
}

# collecting data
$enc = [System.Text.Encoding]::UTF8
$fetched_data = @()
$i = 1
$totrec = $rawdata.rows.Count
$parsebar = ProgressBar
foreach ($record in $rawdata.rows) {
    $fetched_record = @{}
    $fetched_record.FULLNAME = $record.name
    $fetched_record.USRNAME = $record.username
    $fetched_record.EMAIL = $record.email
    $fetched_record.PHONE = $record.phone
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
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-EstrazioneUtenti.csv'


$i = 1
$totrec = $fetched_data.Count
$parsebar = ProgressBar
foreach ($item in $fetched_data) {
    $string = ("AGM{0:d5};{1};{2};{3};{4}" -f ($i,$item.FULLNAME,$item.USRNAME,$item.EMAIL,$item.PHONE))
    $string = $string -replace ';_;', ';NULL;'
    $string = $string -replace ';\s*;', ';NULL;'
    $string = $string -replace ';\s*$', ';NULL'
    $string = $string -replace '&#039;', "'"
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

