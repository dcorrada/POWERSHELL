<#
Name......: WinBuild.ps1
Version...: 25.03.1
Author....: Dario CORRADA

Questo script interroga il DB AGMSkyline e recupera le versioni di build di 
tutti gli asset inventariati
#>


<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
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

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module -Name "$workdir\Modules\Forms.psm1"
        Import-Module SimplySql
        Import-Module ImportExcel
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'SimplySql')) {
            Install-Module SimplySql -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [SimplySql] module: click Ok restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } elseif (!(((Get-InstalledModule).Name) -contains 'ImportExcel')) {
            Install-Module ImportExcel -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [ImportExcel] module: click Ok restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } else {
            [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error') > $null
            Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
            exit
        }
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'


<# *******************************************************************************
                                    ACCESSING
******************************************************************************* #>
# local IP address of the MySQL server
$ahost = '192.168.20.205'

Write-Host -NoNewline "Credential management... "
$pswout = PowerShell.exe -file "$workdir\Safety\Stargate.ps1" -ascript 'AGMskyline'
if ($pswout.Count -eq 2) {
    $MySQLpwd = ConvertTo-SecureString $pswout[1] -AsPlainText -Force    
    $MySQLlogin = New-Object System.Management.Automation.PSCredential($pswout[0], $MySQLpwd)
    Write-Host -ForegroundColor Green ' Ok'
} else {
    [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

# apro una connessione sul DB
Write-Host -NoNewline "Connecting to AGMskyline... "
$ErrorActionPreference= 'Stop'
try {
    Open-MySqlConnection -Server $ahost -Database 'AGMskyline' -Credential $MySQLlogin
    Write-Host -ForegroundColor Green 'DONE'
} catch {
    Write-Host -ForegroundColor Red 'FAILED'
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}
$ErrorActionPreference= 'Inquire'

<# *******************************************************************************
                                    FETCHING
******************************************************************************* #>
Write-Host -NoNewline "Querying... "
$ErrorActionPreference= 'Stop'
try {
    $aquery = @"
SELECT DISTINCT EstrazioneUtenti.FULLNAME, EstrazioneUtenti.EMAIL,
       EstrazioneAsset.HOSTNAME,
       ADcomputers.OS AS 'AD_OS',
       CONCAT(AzureDevices.OSTYPE, ' ', AzureDevices.OSVER) AS 'AZURE_OS',
       TrendMicroparsed.OS AS 'TM_OS'
FROM Xhosts
LEFT JOIN EstrazioneAsset ON Xhosts.ESTRAZIONEASSET = EstrazioneAsset.ID
LEFT JOIN TrendMicroparsed ON Xhosts.TRENDMICROPARSED = TrendMicroparsed.ID
LEFT JOIN ADcomputers ON Xhosts.ADCOMPUTERS = ADcomputers.ID
LEFT JOIN AzureDevices ON Xhosts.AZUREDEVICES = AzureDevices.ID
LEFT JOIN EstrazioneUtenti ON EstrazioneAsset.USRNAME = EstrazioneUtenti.USRNAME
WHERE EstrazioneAsset.STATUS = 'Assegnato'
"@
    $RawData = Invoke-SqlQuery -Query $aquery
    Write-Host -ForegroundColor Green 'DONE'
} catch {
    Write-Host -ForegroundColor Red 'FAILED'
    Write-Output "`nError: $($error[0].ToString())"
    Pause
}
$ErrorActionPreference= 'Inquire'

# chiudo la connessione al DB
Close-SqlConnection

<# *******************************************************************************
                                    PARSING
******************************************************************************* #>
Write-Host -NoNewline "Parsing..."
$BuildDict = @{
    '18362' = '19H1'
    '18363' = '19H2'
    '19041' = '20H1'
    '19042' = '20H2'
    '19043' = '21H1'
    '19044' = '21H2'
    '19045' = '22H2'
    '22621' = '22H2'
    '22631' = '23H2'
    '26100' = '24H2'
}

$ParsedData = $RawData | ForEach-Object {
    Write-Host -NoNewline '.'
    if (([string]::IsNullOrEmpty($_.AD_OS)) -and ([string]::IsNullOrEmpty($_.AZURE_OS))) {
        # no data available (no joined to domain and/or other OS installed)
        $daBuild = 'na'
    } else {
        $OSs = @()
        foreach ($currentString in ($_.AD_OS, $_.AZURE_OS, $_.TM_OS)) {
            $gotcha = $currentString | Select-String -Pattern ($($BuildDict.Keys))
            if ($gotcha.Matches.Success) {
                $OSs += $BuildDict[$gotcha.Matches.Value]
            }
        }
        if ($OSs.Count -gt 1) {
            $OSs = $OSs | Sort-Object -Descending
            $daBuild = $OSs[0]
        } elseif ($OSs.Count -eq 1) {
            $daBuild = $OSs
        } else {
            $daBuild = 'na'
        }
    }
    New-Object -TypeName PSObject -Property @{
        FULLNAME   = "$($_.FULLNAME)"
        EMAIL      = "$($_.EMAIL)"
        HOSTNAME   = "$($_.HOSTNAME)"
        BUILD      = "$daBuild"
    } | Select FULLNAME, EMAIL, HOSTNAME, BUILD
}
Write-Host -ForegroundColor Green ' DONE'

<# *******************************************************************************
                                    OUTPUT
******************************************************************************* #>
$xlsx_file = "C:$env:HOMEPATH\Downloads\WinBuild-" + (Get-Date -format "yyMMddHHmmSS") + '.xlsx'
$XlsPkg = Open-ExcelPackage -Path $xlsx_file -Create
$XlsPkg = $ParsedData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName 'PrettyFly' -TableName 'PrettyFly' -TableStyle 'Medium2' -AutoSize -PassThru
Close-ExcelPackage -ExcelPackage $XlsPkg -Show