<#
In questa versione dello script faccio un controllo incrociato sul tenant per 
verificare se l'utenza di dominio abbia un corrispettivo UPN.
Questo script produce una tabella con una colonna aggiuntiva "IN_SYNC" di tipo 
enum('Yes', 'No').

Con l'introduzione dell'MFA sugli account Microsoft questa informazione non è 
più necessaria. Mi serve sapere solamente lo status della password per ciò che 
riguarda la sola utenza di dominio su AD
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\miscellaneous\\AGMskyline\\PwdExpire\.ps1$" > $null
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

# import Active Directory module
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory }

# import the MSOnline module
$ErrorActionPreference= 'Stop'
try {
    Import-Module MSOnline
} catch {
    Install-Module MSOnline -Confirm:$False -Force
    Import-Module MSOnline
}
$ErrorActionPreference= 'Inquire'

# connect to Tenant
$ErrorActionPreference= 'Stop'
Try {
    Connect-MsolService
    $ErrorActionPreference= 'Inquire'
}
Catch {
    Write-Host -ForegroundColor Red "*** ERROR ACCESSING TENANT ***"
    # Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}

Write-Host -NoNewline "Retrieving users list..."
Write-Host -NoNewline "."
$o365usrs = Get-MsolUser -All | Where-Object { $_.isLicensed -eq "TRUE" } | select -ExpandProperty 'UserPrincipalName'
Write-Host -NoNewline "."
$user_list = Get-ADUser -Filter * -Property *
Write-Host -NoNewline "."
$esprimo = Get-ADUser -filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties "DisplayName","msDS-UserPasswordExpiryTimeComputed"
$esprimo = $esprimo | Select-Object -Property "DisplayName",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}
$espresso = @{}
Write-Host -NoNewline "."
foreach ($item in $esprimo) {
    $bestbefore = $item.ExpiryDate | Get-Date -Format "yyyy-MM-dd"
    if ($bestbefore -le "1601-01-01") {
        # nel log di PALMA questa data assurda risulta con questo messaggio
        # "warning, account password for 'foobar.baz' is already expired"
        $espresso[$item.DisplayName] = ''
    } else {
        $espresso[$item.DisplayName] = $bestbefore
    }
}
Write-Host -ForegroundColor Green 'DONE'

$rawdata = @{}
$i = 1
$totrec = $user_list.Count
$parsebar = ProgressBar
$now = Get-Date
foreach ($auser in $user_list) {
    $fullname = $auser.Name
    if ($auser.UserPrincipalName -match "^([a-zA-Z]+\.[a-zA-Z]+)@agmsolutions\.net$" ) {
        $usrname = $matches[1]
    } else {
        $usrname = $auser.SamAccountName
    }
    $ErrorActionPreference= 'Stop'
    try {
        $lastpwdset = $auser.PasswordLastSet | Get-Date -Format "yyyy-MM-dd"
    } catch {
        $lastpwdset = 'NULL'
    }
    $ErrorActionPreference= 'Inquire'

    # lo username ha nomenclatura "nome.cognome" ed una licenza o365 assegnata
    if ($usrname -match "^[a-zA-Z]+\.[a-zA-Z]+$") {
        $ErrorActionPreference= 'Stop'
        try {
            $account_expdate = $auser.AccountExpirationDate | Get-Date -Format "yyyy-MM-dd"
        } catch {
            $account_expdate = ''
        }
        $ErrorActionPreference= 'Inquire'
        if ($auser.PasswordNeverExpires -eq 'True') {
            $pwdExpired = 'NeverExpire'
        } else {
            $pwdExpired = $auser.PasswordExpired
        }
        
        if ($espresso.ContainsKey($fullname)) {
            $pwdExpireDate = $espresso[$fullname]
        } else {
            $pwdExpireDate = ''
        }

        if ($o365usrs -contains ("$usrname" + '@agmsolutions.net')) {
            $azuread = 'YES'
        } else {
            $azuread = 'NO'
        }

        $rawdata.$usrname = @{
            USRNAME         = $usrname
            FULLNAME        = $fullname
            ACCOUNT_EXPDATE = $account_expdate
            PWD_LASTSET     = $lastpwdset
            PWD_EXPIRED     = $pwdExpired
            PWD_EXPDATE     = $pwdExpireDate
            IN_SYNC         = $azuread
        }
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
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-PwdExpire.csv'


$i = 1
$totrec = $rawdata.Count
$parsebar = ProgressBar
foreach ($usr in $rawdata.Keys) {
    $new_record = @(
        #$usr,
        $rawdata.$usr.USRNAME,
        $rawdata.$usr.FULLNAME,
        $rawdata.$usr.ACCOUNT_EXPDATE,
        $rawdata.$usr.PWD_LASTSET,
        $rawdata.$usr.PWD_EXPIRED,
        $rawdata.$usr.PWD_EXPDATE,
        $rawdata.$usr.IN_SYNC
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
