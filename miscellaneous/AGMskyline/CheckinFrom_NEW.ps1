<#
1) Lo script vecchio eseguiva la seguente join da un DB backuppato di SnipeIT, 
un dump che ora l'interfaccia web non riesce più a produrre:

SELECT action_logs.created_at AS 'timestamp', action_logs.action_type,
       CONCAT(users.first_name, ' ', users.last_name) AS 'fullname', users.email, users.deleted_at,
       assets.name AS 'asset', assets.serial, status_labels.name AS 'status', assets.notes
FROM action_logs
INNER JOIN users ON action_logs.target_id = users.id
INNER JOIN assets ON action_logs.item_id = assets.id
LEFT JOIN status_labels ON assets.status_id = status_labels.id
WHERE action_logs.action_type IN ('checkin from', 'checkout') AND action_logs.item_type LIKE '%Asset'


2) Il file csv in output dovrebbe mantenere il seguente formato:

"timestamp","action_type","fullname","email","deleted_at","asset","serial","status","notes"
"2019-11-18 10:57:22","checkout","Gianfranco Di Tommaso","gianfranco.ditommaso@agmsolutions.net","2021-04-02 14:15:41","2FRB6Z2","2FRB6Z2","Assegnato","ex Salvatore Gabrieli"
"2019-11-18 11:32:11","checkin from","Gianfranco Di Tommaso","gianfranco.ditommaso@agmsolutions.net","2021-04-02 14:15:41","2FRB6Z2","2FRB6Z2","Assegnato","ex Salvatore Gabrieli"
"2019-11-18 15:32:37","checkout","Andrea Accolla","andrea.accolla@agmsolutions.net",NULL,"6C9D6Z2-TO","6C9D6Z2","da Assegnare","ex Christian Cammarata"
"2019-11-18 15:36:40","checkout","Gianfranco Di Tommaso","gianfranco.ditommaso@agmsolutions.net","2021-04-02 14:15:41","2FRB6Z2","2FRB6Z2","Assegnato","ex Salvatore Gabrieli"
[...]



3) Guardare alla tabella da caricare sul DB: la voce "deleted_at" serve per 
discriminare tra ASSUNTO o CESSATO. Provare, nel nuovo script, a vedere se si 
può fare la stessa cosa incrociando con la hash table degli utenti (quelli 
non presenti dovrebbero essere i cessati).
La tabella sul DB si presenta così:

"ID";"UPTIME";"CHECKINOUT";"FULLNAME";"MAIL";"USR_STATUS";"HOSTNAME";"SERIAL";"ASSET_STATUS"
"AGM00001";"2019-11-18";"checkout";"Gianfranco Di Tommaso";"gianfranco.ditommaso@agmsolutions.net";"CESSATO";"2FRB6Z2";"2FRB6Z2";"Pezzi di ricambio"
"AGM00002";"2019-11-18";"checkin from";"Gianfranco Di Tommaso";"gianfranco.ditommaso@agmsolutions.net";"CESSATO";"2FRB6Z2";"2FRB6Z2";"Pezzi di ricambio"
"AGM00003";"2019-11-18";"checkout";"Andrea Accolla";"andrea.accolla@agmsolutions.net";"ASSUNTO";"6C9D6Z2-TO";"6C9D6Z2";"Assegnato"
"AGM00004";"2019-11-18";"checkout";"Gianfranco Di Tommaso";"gianfranco.ditommaso@agmsolutions.net";"CESSATO";"2FRB6Z2";"2FRB6Z2";"Pezzi di ricambio"
"AGM00005";"2019-11-18";"checkin from";"Andrea Accolla";"andrea.accolla@agmsolutions.net";"ASSUNTO";"6C9D6Z2-TO";"6C9D6Z2";"Assegnato"
[...]



Per cercare tabelle e sintassi delle API guardare su:
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
foreach ($record in $rawdata.rows) {
    Write-Host -NoNewline '.'
    if ($record.action_type -contains ('checkin from', 'checkout')) {
        <# Action to perform if the condition is true #>
    }
    $TheUsers[$record.id] = @{
        TIMESTAMP       = "$($record.created_at.datetime)"
        ACTION_TYPE     = $record.action_type
        USER            = $record.target.name
    } 
}
Write-Host -ForegroundColor Green " DONE"

<# esempio di record
id                 : 8003
icon               : fa fa-barcode
file               :
item               : @{id=314; name=CND8154ZW5 (CND8154ZW5); type=asset}
location           :
created_at         : @{datetime=2025-06-04 10:49:24; formatted=2025-06-04 10:49AM}
updated_at         : @{datetime=2025-06-04 10:49:24; formatted=2025-06-04 10:49AM}
next_audit_date    : @{date=2025-06-04; formatted=2025-06-04}
days_to_next_audit : 1
action_type        : checkin from
admin              : @{id=513; name=Marco Motta; first_name=Marco; last_name=Motta}
target             : @{id=443; name=Pamela Casullo; type=user}
note               : Allestito PC SOSTITUTIVO
signature_file     :
log_meta           :
#>