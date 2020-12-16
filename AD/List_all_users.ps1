<#
Name......: List_all_users.ps1
Version...: 19.08.1
Author....: Dario CORRADA

Questo script accede ad Active Directory ed estrae in un file CSV l'elenco di tutti gli account presenti
#>


# header
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

Import-Module -Name '\\192.168.2.251\Dario\SCRIPT\Moduli_PowerShell\Forms.psm1'

# setto le policy di esecuzione degli script
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'

# Controllo accesso
$login = LoginWindow

# Importo il modulo di Active Directory
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory }

# recupero la lista di tutti gli account
Write-Host "Recupero la lista di tutti gli account..."
$user_list = Get-ADUser -Filter * -Property *
Write-Host -ForegroundColor Green "Trovati" $user_list.Count "account"

$rawdata = @{}
$i = 1
foreach ($user_name in $user_list.Name) {
    
    $ErrorActionPreference = 'SilentlyContinue'

    Clear-Host
    Write-Host "Registrazione" $i "di" $user_list.Count

    $infouser = Get-ADUser -Identity $user_name -Properties *
    $infouser.CanonicalName -match "/(.+)/$user_name$" > $null
    $ou = $matches[1]

    $fullname = $infouser.DisplayName -replace ', ', ','

    if (!($user_name -match "HealthMailbox")) {
        $rawdata.$user_name = @{  
            OrganizationalUnit = $ou         
            Company = $infouser.Company
            Created = $infouser.Created
            Department = $infouser.Department
            Description = $infouser.Description
            DisplayName = $fullname
            EmailAddress = $infouser.EmailAddress
            LastLogonDate = $infouser.LastLogonDate
            Office = $infouser.Office
            Title = $infouser.Title
        }
    }
    $i++
}

$ErrorActionPreference= 'Inquire'

$outfile = "C:\Users\$env:USERNAME\Desktop\AD_Users.csv"
"UserName;OrganizationalUnit;Company;Created;Department;Description;DisplayName;EmailAddress;LastLogonDate;Office;Title" | Out-File $outfile -Encoding ASCII -Append
foreach ($usr in $rawdata.Keys) {
    $uname = $usr.ToLower()
    $new_record = @(
        $uname,
        $rawdata.$usr.OrganizationalUnit,
        $rawdata.$usr.Company,
        $rawdata.$usr.Created,
        $rawdata.$usr.Department,
        $rawdata.$usr.Description,
        $rawdata.$usr.DisplayName,
        $rawdata.$usr.EmailAddress,
        $rawdata.$usr.LastLogonDate,
        $rawdata.$usr.Office,
        $rawdata.$usr.Title
    )
    $new_string = [system.String]::Join(";", $new_record)
    $new_string | Out-File $outfile -Encoding ASCII -Append
}