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
Import-Module -Name "$workdir\Modules\Forms.psm1"

# local IP address of the MySQL server
$ahost = '192.168.1.30'

# MySQL login dialog
[System.Windows.MessageBox]::Show("Insert credentials for MySQL connection",'MYSQL','Ok','Info') | Out-Null
$pswout = PowerShell.exe -file "$workdir\Safety\Stargate.ps1" -ascript 'AGMskyline'
if ($pswout.Count -eq 2) {
    ($dumpUsr, $dumpPwd) = ($pswout[0], $pswout[1])
    $MySQLpwd = ConvertTo-SecureString $pswout[1] -AsPlainText -Force    
    $MySQLlogin = New-Object System.Management.Automation.PSCredential($pswout[0], $MySQLpwd)
} else {
    [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
    exit
}

# Dump del DB
$answ = [System.Windows.MessageBox]::Show("Dump the current DB?",'BACKUP','YesNo','Info')
if ($answ -eq "Yes") {
    # import the Posh-SSH module
    $ErrorActionPreference= 'Stop'
    try {
        Import-Module Posh-SSH
    } catch {
        Install-Module Posh-SSH -Confirm:$False -Force
        Import-Module Posh-SSH
    }
    $ErrorActionPreference= 'Inquire'

    # SSH login dialog
    [System.Windows.MessageBox]::Show("Insert credentials for SSH connection",'SSH','Ok','Info') | Out-Null
    $pswout = PowerShell.exe -file "$workdir\Safety\Stargate.ps1" -ascript 'AGMskyline'
    if ($pswout.Count -eq 2) {
        $SSHpwd = ConvertTo-SecureString $pswout[1] -AsPlainText -Force
        $SSHlogin = New-Object System.Management.Automation.PSCredential($pswout[0], $SSHpwd)
    } else {
        [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
        exit
    }

    # eseguo il dump
    $ErrorActionPreference= 'Stop'
    try {
        $sshsession = New-SSHSession -ComputerName $ahost -Credential $SSHlogin -Force -Verbose
        [string]$backupcmd = ("mysqldump -u {0} -p{1} AGMskyline > /home/dario.corrada/share/{2}-AGMskyline_dump.sql" -f ($dumpUsr, $dumpPwd, (Get-Date -format "yyMMdd")))
        $stdout = Invoke-SSHCommand -SSHSession $sshsession -Command $backupcmd
        Remove-SSHSession -SSHSession $sshsession
    } catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
    }
    $ErrorActionPreference= 'Inquire'
}

# seleziono la cartella dove sono presenti i file csv dei dati da caricare
[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") > $null
$foldername = New-Object System.Windows.Forms.FolderBrowserDialog
$foldername.RootFolder = "MyComputer"
$foldername.ShowNewFolderButton = $false
$foldername.Description = "Seleziona la cartella dove sono `npresenti i file csv da caricare"
$foldername.ShowDialog() > $null
$sharepath = $foldername.SelectedPath

if ((Get-ChildItem -Path ($sharepath + '\*.csv')).Count -le 0) {
    Write-Host -NoNewline "No csv file found..."
    Write-Host -ForegroundColor Red 'ABORTING'
    Pause
    exit
}

# import the SimplySQL module
$ErrorActionPreference= 'Stop'
try {
    Import-Module SimplySql
} catch {
    Install-Module SimplySql -Confirm:$False -Force
    Import-Module SimplySql
}
$ErrorActionPreference= 'Inquire'

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

# raccolgo l'elenco delle tabelle presenti sul DB
$tablelist = Invoke-SqlQuery -Query 'SHOW TABLES'

# recupero la lista dei file e parso nonme tabella e data
foreach ($afile in (Get-ChildItem -Path ($sharepath + '\*.csv') -Name)) {
    $matches = @()
    $ErrorActionPreference= 'Stop'
    try {
        $afile -match "^([0-9]+)\-([a-zA-Z_\-\.\\\s0-9:]+)\.csv$" > $null
        $uptime = [Datetime]::ParseExact($matches[1], 'yyMMdd', $null) | Get-Date -format "yyyy-MM-dd"
        $tablename = $matches[2]
        if ($tablelist.Tables_in_AGMskyline -contains $tablename) {
            Write-Host -NoNewline -ForegroundColor Yellow ("Loading data for table [{0}] updated on {1}..." -f ($tablename, $uptime))
            
            $ErrorActionPreference= 'Stop'
            try {
                # svuoto la tabella dei vecchi record
                Invoke-SqlQuery -Query "TRUNCATE $tablename"
                Write-Host -NoNewline -ForegroundColor Yellow '.'
                
                # importo il csv di dati grezzi
                Invoke-SqlQuery -Query "LOCK TABLES $tablename WRITE"
                Write-Host -NoNewline -ForegroundColor Yellow '.'
                $rawdata = Get-Content -Path ($sharepath + '\' + $afile)
                $totrec = $rawdata.Count
                $parsebar = ProgressBar
                for ($i = 0; $i -lt $rawdata.Count; $i++) {
                    # il limite di $k indica il numero di valori da inserire per ogni query di INSERT
                    for ($k = 0; $k -lt 50; $k++) {
                        if ($i -lt $rawdata.Count) {
                            $rawvalue = $rawdata[$i] -replace "'", "\'" # escaping single quote, see https://stackoverflow.com/questions/4803354/how-do-i-insert-a-special-character-such-as-into-mysql
                            $rawvalue = $rawvalue -replace '"', "'"
                            $rawvalue = $rawvalue -replace ';', ", "
                            if ($k -eq 0) {
                                $QueryString = "INSERT INTO $tablename VALUES ($rawvalue)"
                            } else {
                                $QueryString = $QueryString + ",($rawvalue)"
                            }
                            $i++
                        }
                    }
                    $i--
                    Invoke-SqlQuery -Query "$QueryString"
                    Start-Sleep -Milliseconds 100
                    <# *** BUFFERING ***
                        Inserisco uno sleep per permettere il triggering dell'autocommit.
                        Sull'argomento qui c'è la pagina di riferimento di MySQL:

                        https://dev.mysql.com/doc/refman/8.0/en/innodb-autocommit-commit-rollback.html
                    #>
                    $QueryString = $null
                    $rawvalue = $null

                    # progress
                    $percent = (($i) / $totrec)*100
                    if ($percent -gt 100) {
                        $percent = 100
                    }
                    $formattato = '{0:0.0}' -f $percent
                    [int32]$progress = $percent   
                    $parsebar[2].Text = ("Add {0} out of {1} recs [{2}%] on {3}" -f (($i-1), $totrec, $formattato, $tablename))
                    if ($progress -ge 100) {
                        $parsebar[1].Value = 100
                    } else {
                        $parsebar[1].Value = $progress
                    }
                    [System.Windows.Forms.Application]::DoEvents()
                }
                $parsebar[0].Close()
                Write-Host -NoNewline -ForegroundColor Yellow '.'
                Invoke-SqlQuery -Query "UNLOCK TABLES"
                Write-Host -NoNewline -ForegroundColor Yellow '.'

                # aggiornamento update su UpdatedTables
                Invoke-SqlQuery -Query "UPDATE `UpdatedTables` SET `UPDATED` = @atime WHERE `UpdatedTables`.`ATABLE` = @atable" -Parameters @{atime = $uptime.ToString(); atable = $tablename}
                Write-Host -NoNewline -ForegroundColor Yellow '.'

                Write-Host -ForegroundColor Green 'DONE'
            } catch {
                Write-Host -ForegroundColor Red 'FAILED'
                Write-Output "`nError: $($error[0].ToString())"
                if ($rawvalue -ne $null) {
                    # faccio un check dell'eventuale riga del csv che possa aver sollevato un'eccezione durante una INSERT
                    Write-Host -ForegroundColor Blue "`n>>>$rawvalue<<<"
                    Pause
                    Invoke-SqlQuery -Query "UNLOCK TABLES"
                } else {
                    Pause
                    exit
                }
            }
            $ErrorActionPreference= 'Inquire'
        } else {
            Write-Host -ForegroundColor Red "Table [$tablename] not found on DB"
            Pause
        }
    } catch {
        Write-Host -ForegroundColor Red "Bad filename <$afile>"
        # log x eccezioni non gestite, decommentare x debug
        # Write-Output "Error: $($error[0].ToString())"
        Pause
    }
    $ErrorActionPreference= 'Inquire'
}

# tabelle riferimenti incrociati
$answ = [System.Windows.MessageBox]::Show("Reset crossrefs tables?",'XREF','YesNo','Info')
if ($answ -eq "Yes") {
    Write-Host -NoNewline -ForegroundColor Yellow "Reset table Xhosts..."
    $ErrorActionPreference= 'Stop'
    try {
        Invoke-SqlQuery -Query "DROP TABLE IF EXISTS Xhosts"
        Write-Host -NoNewline -ForegroundColor Yellow '.'

        $aquery = @"
CREATE TABLE Xhosts (ID INT NOT NULL AUTO_INCREMENT, PRIMARY KEY (ID)) ENGINE=InnoDB
SELECT EstrazioneAsset.HOSTNAME AS 'HOSTNAME', EstrazioneAsset.ID AS 'ESTRAZIONEASSET',
        ADcomputers.ID AS 'ADCOMPUTERS', 
        AzureDevices.ID AS 'AZUREDEVICES', 
        GFIparsed.ID AS 'GFIPARSED',
        TrendMicroparsed.ID AS 'TRENDMICROPARSED',
        CheckinFrom.ID AS 'CHECKINFROM'
FROM EstrazioneAsset
LEFT JOIN ADcomputers ON EstrazioneAsset.HOSTNAME = ADcomputers.HOSTNAME
LEFT JOIN AzureDevices ON EstrazioneAsset.HOSTNAME = AzureDevices.HOSTNAME
LEFT JOIN GFIparsed ON EstrazioneAsset.HOSTNAME = GFIparsed.HOSTNAME
LEFT JOIN TrendMicroparsed ON EstrazioneAsset.HOSTNAME = TrendMicroparsed.HOSTNAME
LEFT JOIN CheckinFrom ON EstrazioneAsset.HOSTNAME = CheckinFrom.HOSTNAME        
"@
        Invoke-SqlQuery -Query $aquery
        Write-Host -NoNewline -ForegroundColor Yellow '.'

        # aggiornamento su UpdatedTables
        $tablename = 'Xhosts'
        $uptime = Get-Date -format "yyyy-MM-dd"
        Invoke-SqlQuery -Query "UPDATE `UpdatedTables` SET `UPDATED` = @atime WHERE `UpdatedTables`.`ATABLE` = @atable" -Parameters @{atime = $uptime.ToString(); atable = $tablename}
        Write-Host -NoNewline -ForegroundColor Yellow '.'

        Write-Host -ForegroundColor Green 'DONE'
    } catch {
        Write-Host -ForegroundColor Red 'FAILED'
        Write-Output "`nError: $($error[0].ToString())"
        Pause
    }
    $ErrorActionPreference= 'Inquire'

    Write-Host -NoNewline -ForegroundColor Yellow "Reset table Xusers..."
    $ErrorActionPreference= 'Stop'
    try {
        Invoke-SqlQuery -Query "DROP TABLE IF EXISTS Xusers"
        Write-Host -NoNewline -ForegroundColor Yellow '.'

        $aquery = @"
CREATE TABLE Xusers (ID INT NOT NULL AUTO_INCREMENT, PRIMARY KEY (ID)) ENGINE=InnoDB
SELECT ADusers.FULLNAME AS 'FULLNAME', o365licenses.ID AS 'O365LICENSES',
       AzureDevices.ID AS 'AZUREDEVICES', DLmembers.ID AS 'DLMEMBERS',
       SchedeAssunzione.ID AS 'SCHEDEASSUNZIONE',
       SchedeTelefoni.ID AS 'SCHEDETELEFONI',
       SchedeSIM.ID AS 'SCHEDESIM',
       ThirdPartiesLicenses.ID AS 'THIRDPARTIESLICENSES',
       EstrazioneUtenti.ID as 'ESTRAZIONEUTENTI',
       CheckinFrom.ID AS 'CHECKINFROM',
       ADusers.ID AS 'ADUSERS',
       PwdExpire.ID AS 'PWDEXPIRE'
FROM ADusers
LEFT JOIN o365licenses ON ADusers.UPN = o365licenses.MAIL
LEFT JOIN AzureDevices ON ADusers.UPN = AzureDevices.MAIL
LEFT JOIN DLmembers ON ADusers.UPN = DLmembers.EMAIL
LEFT JOIN SchedeAssunzione ON ADusers.UPN = SchedeAssunzione.MAIL
LEFT JOIN SchedeTelefoni ON LOWER(ADusers.FULLNAME) LIKE CONCAT('%', LOWER(SchedeTelefoni.FULLNAME), '%')
LEFT JOIN SchedeSIM ON LOWER(ADusers.FULLNAME) LIKE CONCAT('%', LOWER(SchedeSIM.FULLNAME), '%')
LEFT JOIN ThirdPartiesLicenses ON ADusers.UPN = ThirdPartiesLicenses.MAIL
LEFT JOIN EstrazioneUtenti ON ADusers.UPN = EstrazioneUtenti.EMAIL
LEFT JOIN CheckinFrom ON ADusers.UPN = CheckinFrom.MAIL
LEFT JOIN PwdExpire ON ADusers.USRNAME = PwdExpire.USRNAME
"@
        Invoke-SqlQuery -Query $aquery
        Write-Host -NoNewline -ForegroundColor Yellow '.'

        # aggiornamento su UpdatedTables
        $tablename = 'Xusers'
        $uptime = Get-Date -format "yyyy-MM-dd"
        Invoke-SqlQuery -Query "UPDATE `UpdatedTables` SET `UPDATED` = @atime WHERE `UpdatedTables`.`ATABLE` = @atable" -Parameters @{atime = $uptime.ToString(); atable = $tablename}
        Write-Host -NoNewline -ForegroundColor Yellow '.'
        Write-Host -ForegroundColor Green 'DONE'
    } catch {
        Write-Host -ForegroundColor Red 'FAILED'
        Write-Output "`nError: $($error[0].ToString())"
        Pause
    }
    $ErrorActionPreference= 'Inquire'
}


# chiudo la connessione al DB
Close-SqlConnection
