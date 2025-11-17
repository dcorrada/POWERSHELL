<#
Name......: UnjoinEmAll.ps1
Version...: 25.10.1
Author....: Dario CORRADA

Script per rimuovere da dominio asset non assegnati
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
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent  | Split-Path -Parent

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module -Name "$workdir\Modules\Forms.psm1"
        Import-Module ImportExcel
        $ThirdParty = 'Ok'
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'Microsoft.Graph')) {
            Install-Module Microsoft.Graph -Scope AllUsers -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [Microsoft.Graph] module: click Ok to restart the script",'RESTART','Ok','warning') > $null
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
                                  SNIPE IT
******************************************************************************* #>
Write-Host -NoNewline -ForegroundColor Cyan "`nFetching SnipeIT unassigned assets"
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

# collecting data
$enc = [System.Text.Encoding]::UTF8
$SnipeIT_data = @()
foreach ($record in $rawdata.rows) {
    $fetched_record = @{}
    $fetched_record.STATUS = $record.status_label.name
    if ($fetched_record.STATUS -cne 'Assegnato') {
        $fetched_record.NAME = $record.name        
        $SnipeIT_data += $fetched_record
        Write-Host -NoNewline '+'
    } else {
        Write-Host -NoNewline '.'
    }
}
Write-Host -NoNewline -ForegroundColor Green " Done`n"
Start-Sleep -Milliseconds 1500

<# *******************************************************************************
                                ACTIVE DIRECTORY
******************************************************************************* #>
Write-Host -NoNewline -ForegroundColor Cyan "`nLooking for assets on AD"
$Joined_assets = @{}
foreach ($item in $SnipeIT_data) {
    $ErrorActionPreference= 'Stop'
    try {
        $infopc = Get-ADComputer -Identity $item.NAME -Properties *
        $infopc.CanonicalName -match "/(.+)/$($item.NAME)$" > $null
        $ou = $matches[1]
        Write-Host -NoNewline '+'
        try {
            $Joined_assets[$item.NAME] = @{
                STATUS  = $item.STATUS
                OU      = $ou 
                DATE    = $infopc.LastLogonDate | Get-Date -format "yyyy-MM-dd"
            }
        }
        catch {
            $Joined_assets[$item.NAME] = @{
                STATUS  = $item.STATUS
                OU      = $ou 
                DATE    = 'na'
            }
        }
    }
    catch {
        Write-Host -NoNewline '.'
    }
    $ErrorActionPreference= 'Inquire'
}
Write-Host -NoNewline -ForegroundColor Green " Done`n"
Start-Sleep -Milliseconds 1500

<# *******************************************************************************
                                GET OUTPUT
******************************************************************************* #>
Write-Host -NoNewline -ForegroundColor Cyan "`nPreparing Excel report"
$xlsx_file = "C:$env:HOMEPATH\Downloads\UnjoinEmAll-" + (Get-Date -format "yyMMddHHmm") + '.xlsx'
$XlsPkg = Open-ExcelPackage -Path $xlsx_file -Create

$ErrorActionPreference= 'Stop'
try {  
    $label = 'UnassignedAssets'
    $inData = $Joined_assets.Keys | Foreach-Object {

        New-Object -TypeName PSObject -Property @{
            HOSTNAME  = "$_"
            STATUS    = "$($Joined_assets[$_].STATUS)"
            OU        = "$($Joined_assets[$_].OU)"
            LASTLOGON = "$($Joined_assets[$_].DATE)"
        } | Select HOSTNAME, OU, STATUS, LASTLOGON
        Write-Host -NoNewline '.'
    }
    $XlsPkg = $inData | Export-Excel -ExcelPackage $XlsPkg -WorksheetName $label -TableName $label -TableStyle 'Medium1' -AutoSize -PassThru
} catch {
    [System.Windows.MessageBox]::Show("Error updating data",'ABORTING','Ok','Error') | Out-Null
    Write-Host -ForegroundColor Red ' FAIL'
    Write-Host -ForegroundColor Yellow "ERROR: $($error[0].ToString())"
    exit
}
$ErrorActionPreference= 'Inquire'

Close-ExcelPackage -ExcelPackage $XlsPkg -Show
Write-Host -NoNewline -ForegroundColor Green " Done`n"
Start-Sleep -Milliseconds 1500

$proceed2ujoin = [System.Windows.MessageBox]::Show(@"
[$xlsx_file] 

Questo file illustra una lista di asset che possono essere 
rimossi da dominio.

Cliccando su "Si" questo file verra' ricaricato dallo script, 
che procedera' alla rimozione di tutti gli asset che legge 
da tale file. Puoi quindi editare e salvare questo file 
prima di procedere. 

Cliccando su "No" lo script termina qui. 
"@,'UNJOIN ASSET','YesNo','Info')


<# *******************************************************************************
                             UNJOIN FROM AD
******************************************************************************* #>
if ($proceed2ujoin -eq 'Yes') {
    [System.Windows.MessageBox]::Show("Prima di procedere chiudere il file, se aperto in Excel`n[$xlsx_file]",'UNJOIN ASSET','Ok','Warning') | Out-Null

    # getting AD credentials
    Write-Host -NoNewline "Credential management... "
    $pswout = PowerShell.exe -file "$workdir\Safety\Stargate.ps1" -ascript 'UnjoinEmAll'
    if ($pswout.Count -eq 2) {
        $ad_login = New-Object System.Management.Automation.PSCredential($pswout[0], (ConvertTo-SecureString $pswout[1] -AsPlainText -Force))
    } else {
        [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
        Write-Host -ForegroundColor Red "Ko"
        Pause
        exit
    }
    Write-Host -ForegroundColor Green 'Ok'

    $DeadList = Import-Excel -Path $xlsx_file -WorksheetName 'UnassignedAssets'
    $StillAlive = @()
    Clear-Host
    Write-Host -ForegroundColor Yellow "*** UNJOINIG HOSTS ***`n"
    foreach ($Zombie in $DeadList) {
        Write-Host -ForegroundColor Blue -NoNewline "Removing $zombie "
        $ErrorActionPreference= 'Stop'
        try {
            Remove-ADComputer -Identity $Zombie.HOSTNAME -Credential $ad_login -confirm:$false
            Write-Host -ForegroundColor Green " Done"
        }
        catch {
            $StillAlive += $Zombie
            Write-Host -ForegroundColor Red "Failed"
        }
        $ErrorActionPreference= 'Inquire'
        Start-Sleep -Milliseconds 500
    }

    if ($StillAlive.Count -gt 0) {
        $xlsx_alive = "C:$env:HOMEPATH\Downloads\NotUnjoined-" + (Get-Date -format "yyMMddHHmm") + '.xlsx'
        $AlivePkg = Open-ExcelPackage -Path $xlsx_alive -Create
        $AlivePkg = $StillAlive | Export-Excel -ExcelPackage $AlivePkg -WorksheetName 'UnassignedAssets' -TableName 'UnassignedAssets' -TableStyle 'Medium1' -AutoSize -PassThru
        
        $answer = [System.Windows.MessageBox]::Show(@"
Alcuni asset non sono stati rimossi da dominio.
Aprire il seguente file?

[$xlsx_alive] 
"@,'UNJOIN FAILED','YesNo','Error')
        if ($answer -eq 'Yes') {
            Close-ExcelPackage -ExcelPackage $AlivePkg -Show
        } else {
            Close-ExcelPackage -ExcelPackage $AlivePkg
        }
    }

    Remove-Item -Path $xlsx_file -Force
}
