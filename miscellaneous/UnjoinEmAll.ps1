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
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module -Name "$workdir\Modules\Forms.psm1"
        Import-Module ActiveDiirectory
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
    }
    Write-Host -NoNewline "."
}
Write-Host -NoNewline -ForegroundColor Green "Done`n"
Start-Sleep -Milliseconds 1500

<# *******************************************************************************
                                ACTIVE DIRECTORY
******************************************************************************* #>
Write-Host -NoNewline -ForegroundColor Cyan "`nLooking for assets on AD"
$Joined_assets = @()
foreach ($item in $SnipeIT_data) {
    $infopc = Get-ADComputer -Identity 'pippo' -Properties *
}
Write-Host -NoNewline -ForegroundColor Green "Done`n"
Start-Sleep -Milliseconds 1500

<# +++ TODO LIST +++

2) Interrogare AD e selezionare solo quegli asset ancora a dominio

3) Interrogare AzureAD e raccogliere i last logon di questa selezione

4) Produrre un Excel di questa selezione e mostrarlo: mettere in pausa lo script

5) Se confermato, eliminare da AD questi host (usare le credenziali "adm.nome.cognome")
#>