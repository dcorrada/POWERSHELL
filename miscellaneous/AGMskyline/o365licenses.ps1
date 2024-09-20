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

# retrieve credentials
Write-Host -NoNewline "Credential management... "
$pswout = PowerShell.exe -file "$workdir\Graph\AppKeyring.ps1"
if ($pswout.Count -eq 4) {
    $UPN = $pswout[0]
    $clientID = $pswout[1]
    $tenantID = $pswout[2]
    Write-Host -ForegroundColor Green ' Ok'
} else {
    [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

# connect to Tenant
Write-Host -NoNewline "Connecting to the Tenant..."
$ErrorActionPreference= 'Stop'
Try {
    $splash = Connect-MgGraph -ClientId $clientID -TenantId $tenantID 
    Write-Host -ForegroundColor Green ' Ok'
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("Error connecting to the Tenant",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

# retrieve all users that are licensed
$Users = Get-MgUser -All -Property UserPrincipalName, DisplayName, CreatedDateTime, assignedLicenses `
                    -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" `
                    -ConsistencyLevel eventual -CountVariable Records `
        | Select-Object UserPrincipalName, DisplayName, CreatedDateTime, assignedLicenses `
        | Sort-Object DisplayName

# SkuID to SkuPartNumber hash table
$SkuID2name = @{}
Get-MgSubscribedSku | Select -Property SkuId, SkuPartNumber | foreach { 
    $SkuID2name[$_.SkuId] = $_.SkuPartNumber 
}

# initialize dataframe for collecting data
$parseddata = @{}

$tot = $Users.Count
$usrcount = 0
$parsebar = ProgressBar
Write-Host -NoNewline "Collecting Data..."
foreach ($User in $Users) {
    $usrcount ++

    $username = $User.UserPrincipalName
    $fullname = $User.DisplayName
    $started = $User.CreatedDateTime | Get-Date -format "yyyy/MM/dd"

    $licenses = $User.assignedLicenses
    $parseddata[$username] = @{
        'fullname' = 'NULL'
        'email' = 'NULL'
        'licenza' = @()
        'start' = 'NULL'
    }
    if ($licenses.Count -ge 1) { # at least one license
        foreach ($license in $licenses) {
            $licenseName = $SkuID2name[$license.SkuID]
            $parseddata[$username].fullname = $fullname
            $parseddata[$username].email = $username
            $parseddata[$username].licenza += $licenseName
            $parseddata[$username].start = $started
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

# disconnect from Tenant
$infoLogout = Disconnect-Graph

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

