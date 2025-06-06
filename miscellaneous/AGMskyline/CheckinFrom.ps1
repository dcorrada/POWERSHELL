<#
Per cercare tabelle e sintassi delle API Resttful guardare su:
https://snipe-it.readme.io/reference/api-overview
#>

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

# settaggi per accesso a SnipeIT
$uri_prefix = 'http://192.168.2.184/'
$token_file = $env:LOCALAPPDATA + '\SnipeIT.token'
$token_string = Get-Content $token_file
$headers = @{
    'Authorization' = "Bearer $token_string"
    'Accept' = 'application/json'
    'Content-Type' = 'application/json'
}

# query tabella assets
Write-Host -NoNewline "Fetching asset list..."
$query_params = @(
    'limit=100000'
    'offset=0'
)
$uri_suffix = '?' + ($query_params -join '&')
$uri = $uri_prefix + 'api/v1/hardware' + $uri_suffix
$ErrorActionPreference= 'Stop'
Try {
    $rawdata = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("$($error[0].ToString())`n`nPlease check whenever token has not been expired",'ABORTING','Ok','Error') | Out-Null
    exit
}
$TheAssets = @{}
foreach ($record in $rawdata.rows) {
    Write-Host -NoNewline '.'
    $TheAssets[$record.name] = @{
        ASSET       = $record.name
        SERIAL      = $record.serial
        STATUS      = $record.status_label.name
    } 
}
Write-Host -ForegroundColor Green " DONE"

# query tabella users
Write-Host -NoNewline "Fetching users list..."
$query_params = @(
    'limit=100000'
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
    [System.Windows.MessageBox]::Show("$($error[0].ToString())`n`nPlease check whenever token has not been expired",'ABORTING','Ok','Error') | Out-Null
    exit
}
$TheUsers = @{}
foreach ($record in $rawdata.rows) {
    Write-Host -NoNewline '.'
    $TheUsers[$record.name] = @{
        FULLNAME    = $record.name
        EMAIL       = $record.email
    } 
}
Write-Host -ForegroundColor Green " DONE"

# query tabella logs
Write-Host -NoNewline "Fetching activity logs..."
$query_params = @(
    'limit=100000'
    'offset=0'
)
$uri_suffix = '?' + ($query_params -join '&')
$uri = $uri_prefix + 'api/v1/reports/activity' + $uri_suffix
$ErrorActionPreference= 'Stop'
Try {
    $rawdata = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("$($error[0].ToString())`n`nPlease check whenever token has not been expired",'ABORTING','Ok','Error') | Out-Null
    exit
}
$TheLogs = @{}
$BlackList = @('Sim Telefonica (3401730777)') # lista di asset e/o refusi da non considerare
foreach ($record in $rawdata.rows) {
    if (($record.item.type -eq 'asset') -and (('checkin from', 'checkout') -contains $record.action_type) -and ($BlackList -cnotcontains $record.item.name)) {
        $record.item.name -match "^([A-Z0-9\-]+)\s\(" | Out-Null
        $AssetFound = "$($matches[1])"
        if ($AssetFound -eq $null) {
            [System.Windows.MessageBox]::Show("Unexpected asset found`"$($record.item.name)`"`nUpdate `$Blacklist array, if necessary, and rerun the script",'ABORTING','Ok','Warning') | Out-Null
            Write-Host "$($record.item.name)"
            Exit
        } elseif ($TheAssets.ContainsKey($AssetFound)) {
            Write-Host -NoNewline '.'
            $UsrAlive = 'CESSATO'
            $UsrMail = 'NULL'

            # BUG: x qualche strano motivo, su alcune transazioni, questo dato e' saltato (v. asset 5CG9014LPR sul checkout del 2023-06-12 2:14PM)
            if ($record.target.name -eq $null) {
                $TargetUsr = 'unknown'
            } else {
                $TargetUsr = $record.target.name
            }

            if ($TheUsers.ContainsKey($TargetUsr)) {
                $UsrAlive = 'ASSUNTO'
                $UsrMail = $TheUsers[$TargetUsr].EMAIL
                if ($UsrMail -notmatch '@') {
                    $UsrMail = 'NULL'
                }
            }

            $TheLogs[$record.id] = @{
                UPTIME          = "$($record.created_at.datetime)"
                CHECKINOUT      = $record.action_type
                FULLNAME        = $TargetUsr
                MAIL            = $UsrMail
                USRSTATUS       = $usrAlive
                HOSTNAME        = $TheAssets[$AssetFound].ASSET
                SERIAL          = $TheAssets[$AssetFound].SERIAL
                ASSET_STATUS    = $TheAssets[$AssetFound].STATUS
            }
        }
    }
}
Write-Host -ForegroundColor Green " DONE"

# writing output file
Write-Host -NoNewline "Writing output file... "
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-CheckinFrom.csv'

$i = 1
$totrec = $TheLogs.Keys.Count
$parsebar = ProgressBar
foreach ($item in $TheLogs.Keys) {
    $string = ("AGM{0:d5};{1};{2};{3};{4};{5};{6};{7};{8}" -f ( `
        $i, ` 
        $TheLogs[$item].UPTIME, `
        $TheLogs[$item].CHECKINOUT, `
        $TheLogs[$item].FULLNAME, `
        $TheLogs[$item].MAIL,`
        $TheLogs[$item].USRSTATUS, `
        $TheLogs[$item].HOSTNAME, `
        $TheLogs[$item].SERIAL, `
        $TheLogs[$item].ASSET_STATUS
    ))
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