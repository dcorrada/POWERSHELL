<#
Name......: PSWallet.ps1
Version...: 24.05.a
Author....: Dario CORRADA

PSWallet aims to be the credential manager tool in order to handle the various 
login attempts required alongside the scripts of this git repository. 
It will store, fetch and update credential onto a SQLite database.
#>


<# !!! TODO LIST !!!

1) Implementare e testare cmdlet per interfacciarsi su SQLite, vedi refs:
    * https://sqldocs.org/sqlite/sqlite-with-powershell/
    * https://github.com/RamblingCookieMonster/PSSQLite
    * https://www.powershellgallery.com/packages/PSSQLite/1.1.0

2) La prima tabella da creare sara' un log degli accessi che traccera':
    * timestamp    
    * username
    * hostname
    * script che invoca il wallet
    * tipologia di azione (lettura, edit, indel, ...)

3) Determinare i flussi IO sul wallet:
    * in input dagli script che lo invocano 
      (vedi il wrapper usato per recuperare i GET_RAWDATA.ps1 di AGMskyline)
    * in output per raccogliere le info fornite dal wallet
      ( https://stackoverflow.com/questions/8097354/how-do-i-capture-the-output-into-a-variable-from-an-external-process-in-powershe )
    * per ogni script che invoca il wallet...
        + creare una tabella dedicata di credenziali
        + creare un behaviour dedicato (ie su AssignedLicenses.ps1 viene 
          buttata fuori solo una lista di credenziali da scegliere)

4) criptare l'intero file del DB o i singoli record?

5) Versioning: una volta terminato lo sbozzamento dello script rimuoverlo dal 
   branch "tempus" e creare un branch proprio di testing "PSWallet", facendolo 
   ramificare dal branch "unstable". Sul branch "PSWallet" verranno 
   implementate le integrazioni degli script che invocano il wallet.

#>


<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Safety\\PSWallet\.ps1$" > $null
$workdir = $matches[1]
<# for testing purposes
$workdir = Get-Location
$workdir = $workdir.Path
#>

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
try {
    Import-Module -Name "$workdir\Modules\Gordian.psm1"
    Import-Module -Name "$workdir\Modules\Forms.psm1"
    Import-Module PSSQLite
} catch {
    if (!(((Get-InstalledModule).Name) -contains 'PSSQLite')) {
        Install-Module PSSQLite -Confirm:$False -Force
        [System.Windows.MessageBox]::Show("Installed [PSSQLite] module: please restart the script",'RESTART','Ok','warning')
        exit
    } else {
        [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error')
        Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
        exit
    }
}
$ErrorActionPreference= 'Inquire'


<# *******************************************************************************
                                INITIALIZATION
******************************************************************************* #>
$cryptofile = $env:LOCALAPPDATA + '\PSWallet.encrypted'
$dbfile = $env:LOCALAPPDATA + '\PSWallet.sqlite'
do {
    if (!(Test-Path $cryptofile -PathType Leaf)) {
        # create Key file if no DB file exists
        [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Title = "Save Key File"
        $SaveFileDialog.initialDirectory = "$env:LOCALAPPDATA"
        $SaveFileDialog.FileName = 'PSWallet.key'
        $SaveFileDialog.filter = 'Key file (*.key)| *.key'
        $SaveFileDialog.ShowDialog() | Out-Null
        $keyfile = $SaveFileDialog.filename
        CreateKeyFile -keyfile "$keyfile" | Out-Null
    } else {
        [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Title = "Open Key File"
        $OpenFileDialog.initialDirectory = "$env:LOCALAPPDATA"
        $OpenFileDialog.filter = 'Key file (*.key)| *.key'
        $OpenFileDialog.ShowDialog() | Out-Null
        $keyfile = $OpenFileDialog.filename   
    }
} while ([string]::IsNullOrEmpty($keyfile))


Write-Host -NoNewline 'Accessing DB file... '
$ErrorActionPreference= 'Stop'
try {
    if (Test-Path $cryptofile -PathType Leaf) {
        DecryptFile -keyfile "$keyfile" -infile "$cryptofile" | Out-File -FilePath "$dbfile"
        $SQLiteConnection = New-SQLiteConnection -DataSource $dbfile
    } else {
        Write-Host -NoNewline -ForegroundColor Yellow 'No DB found, create it '
        $SQLiteConnection = New-SQLiteConnection -DataSource $dbfile

        # creating DB schema
        Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @'
CREATE TABLE `Logs` (
    `USER` varchar(80),
    `HOST` varchar(80),
    `ACTION` varchar(80),
    `UPTIME` datetime,
    `DESC` text
);
'@
    }

    # login
    Invoke-SqliteQuery -SQLiteConnection $Connection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'login',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@

    Write-Host -ForegroundColor Green 'Ok'
} catch {
    Write-Host -ForegroundColor Red 'Ko'
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Error accessing database",'ABORTING','Ok','Error')
    exit
}
$ErrorActionPreference= 'Inquire'










<# *******************************************************************************
                                CLOSURE
******************************************************************************* #>
Write-Host -NoNewline 'Closing DB file... '
$ErrorActionPreference= 'Stop'
try {
    # logout
    Invoke-SqliteQuery -SQLiteConnection $Connection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'logout',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@

    $SQLiteConnection.Close()
    EncryptFile -keyfile "$keyfile" -infile "$dbfile" -outfile "$cryptofile" | Out-Null
    Write-Host -ForegroundColor Green 'Ok'
} catch {
    Write-Host -ForegroundColor Red 'Ko'
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Error closing database",'ABORTING','Ok','Error')
    exit
}
$ErrorActionPreference= 'Inquire'


