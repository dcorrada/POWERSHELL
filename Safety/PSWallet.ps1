Param([string]$ExtDesc='NULL', 
    [string]$ExtUsr=$env:USERNAME, [string]$ExtHost=$env:COMPUTERNAME,
    [string]$ExtUptime=(Get-Date -format "yyyy-MM-dd HH:mm:ss"), [string]$ExtAction='read')

<#
Name......: PSWallet.ps1
Version...: 24.05.a
Author....: Dario CORRADA

PSWallet aims to be the credential manager tool in order to handle the various 
login attempts required alongside the scripts of this git repository. 
It will store, fetch and update credential onto a SQLite database.

Refs:
* https://github.com/RamblingCookieMonster/PSSQLite
* https://www.powershellgallery.com/packages/PSSQLite/1.1.0

#>


<# !!! TODO LIST !!!

3) Determinare i flussi IO sul wallet:
    * in input dagli script che lo invocano 
      (vedi il wrapper usato per recuperare i GET_RAWDATA.ps1 di AGMskyline)
    * in output per raccogliere le info fornite dal wallet
      ( https://stackoverflow.com/questions/8097354/how-do-i-capture-the-output-into-a-variable-from-an-external-process-in-powershe )
    * per ogni script che invoca il wallet...
        + creare una tabella dedicata di credenziali
        + creare un behaviour dedicato (ie su AssignedLicenses.ps1 viene 
          buttata fuori solo una lista di credenziali da scegliere)

4) Cifrare, on th fly, le password nelle tabelle inserite come SecureString.
   Aggiornare Gordian.psm1 con funzioni per criptare/decriptare testo.

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
do {
    try {
        Import-Module -Name "$workdir\Modules\Forms.psm1"
        Import-Module -Name "$workdir\Modules\Gordian.psm1"
        Import-Module PSSQLite
        Import-Module ImportExcel
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'PSSQLite')) {
            Install-Module PSSQLite -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [PSSQLite] module: please restart the script",'RESTART','Ok','warning') | Out-Null
            $ThirdParty = 'Ko'
        } elseif (!(((Get-InstalledModule).Name) -contains 'ImportExcel')) {
            Install-Module ImportExcel -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [ImportExcel] module: please restart the script",'RESTART','Ok','warning') | Out-Null
            $ThirdParty = 'Ko'
        } else {
            [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error') | Out-Null
            Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
            exit
        }
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                                INITIALIZATION
******************************************************************************* #>
$dbfile = $env:LOCALAPPDATA + '\PSWallet.sqlite'

do {
    if (!(Test-Path $dbfile -PathType Leaf)) {
        # create Key file if no DB file exists
        [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Title = "Create Key File"
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
    if (Test-Path $dbfile -PathType Leaf) {
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
CREATE TABLE `Credits` (
    `USER` varchar(80),
    `PSWD` text,
    `DESC` text
);
'@
    }

    # login
    $SQLiteConnection.Close()
    $SQLiteConnection.Open()
    Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'login',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@
    $SQLiteConnection.Close()
    Write-Host -ForegroundColor Green 'Ok'
} catch {
    Write-Host -ForegroundColor Red 'Ko'
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Error accessing database",'ABORTING','Ok','Error')
    exit
}
$ErrorActionPreference= 'Inquire'


<# *******************************************************************************
                                LOCALES
******************************************************************************* #>
if ($ExtDesc -eq 'NULL') {
    $adialog = FormBase -w 350 -h 170 -text "MAINTENANCE"
    $importCredits = RadioButton -form $adialog -checked $false -x 20 -y 20 -w 500 -h 30 -text "Import Credits table"
    $exportCredits = RadioButton -form $adialog -checked $true -x 20 -y 50 -w 500 -h 30 -text "Export Credits table"
    OKButton -form $adialog -x 100 -y 90 -text "Ok" | Out-Null
    $result = $adialog.ShowDialog()
    if ($importCredits.Checked -eq $true) {
        [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.Title = "Import Table"
        $OpenFileDialog.initialDirectory = "C:$env:HOMEPATH"
        $OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
        $OpenFileDialog.ShowDialog() | Out-Null
        $ImportFile = $OpenFileDialog.filename 
        
        $headers = Compare-Object -ReferenceObject ('USER', 'PSWD', 'DESC') -DifferenceObject ((Import-Excel -Path $ImportFile | Get-Member).Name)
        if ($headers.SideIndicator -contains '<=') {
            [System.Windows.MessageBox]::Show("Required fields doesn't match",'ERROR','Ok','Error')
        } else {
            Write-Host -NoNewline 'Importing...'
            $ImportedTable = Import-Excel -Path $ImportFile | ForEach-Object {
                Write-Host -NoNewline '.'
                New-Object -TypeName PSObject -Property @{
                    USER    = $_.USER
                    PSWD    = EncryptString -keyfile $keyfile -instring $_.PSWD
                    DESC    = $_.DESC
                } | Select USER, PSWD, DESC
            }
            $DataTable = $ImportedTable | Out-DataTable
            # purge any existing record
            $SQLiteConnection.Open()
            Write-Host -NoNewline '.'
            Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query 'DELETE FROM `Credits`;'
            # insert new records
            $stdout = Invoke-SQLiteBulkCopy -DataTable $DataTable -SQLiteConnection $SQLiteConnection -Table Credits -NotifyAfter 0 -ConflictClause Ignore -Verbose -Force
            Write-Host -NoNewline '.'
            Write-Host -ForegroundColor Green ' DONE'
            Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'import Credits',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@
            $SQLiteConnection.Close()
        }
    } elseif ($exportCredits.Checked -eq $true) {
        Write-Host -NoNewline 'Exporting...'
        $SQLiteConnection.Open()
        $rawdata = Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query 'SELECT * FROM `Credits`;'
        Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'export Credits',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@
        $SQLiteConnection.Close()

        $ExportedTable = $rawdata | ForEach-Object {
            Write-Host -NoNewline '.'
            New-Object -TypeName PSObject -Property @{
                USER    = $_.USER
                PSWD    = DecryptString -keyfile $keyfile -instring $_.PSWD
                DESC    = $_.DESC
            } | Select USER, PSWD, DESC
        }
        [System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.Title = "Export Credits"
        $SaveFileDialog.initialDirectory = "C:$env:HOMEPATH\Downloads"
        $SaveFileDialog.FileName = 'PSWallet.xlsx'
        $SaveFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
        $SaveFileDialog.ShowDialog() | Out-Null
        $ExportFile = $SaveFileDialog.filename

        Write-Host -NoNewline '.'
        $XlsPkg = Open-ExcelPackage -Path $ExportFile -Create
        $XlsPkg = $ExportedTable | Export-Excel -ExcelPackage $XlsPkg -WorksheetName 'Credits' -TableName 'Credits' -TableStyle 'Medium3' -AutoSize -PassThru
        Close-ExcelPackage -ExcelPackage $XlsPkg
        Write-Host -ForegroundColor Green ' DONE'
    }
}


<# *******************************************************************************
                                CLOSURE
******************************************************************************* #>
Write-Host -NoNewline 'Closing DB file... '
$ErrorActionPreference= 'Stop'
try {
    # logout
    $SQLiteConnection.Open()
    Invoke-SqliteQuery -SQLiteConnection $SQLiteConnection -Query @"
INSERT INTO Logs (USER, HOST, ACTION, UPTIME) 
VALUES (
    '$($env:USERNAME)',
    '$($env:COMPUTERNAME)',
    'logout',
    '$((Get-Date -format "yyyy-MM-dd HH:mm:ss").ToString())'
);
"@
    $SQLiteConnection.Close()
    Write-Host -ForegroundColor Green 'Ok'
} catch {
    Write-Host -ForegroundColor Red 'Ko'
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    [System.Windows.MessageBox]::Show("Error closing database",'ABORTING','Ok','Error')
    exit
}
$ErrorActionPreference= 'Inquire'


