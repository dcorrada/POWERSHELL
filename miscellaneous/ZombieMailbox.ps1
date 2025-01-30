<#
Name......: ZombieMailbox.ps1
Version...: 25.2.1
Author....: Dario CORRADA

[inserire qui la descrizione]

Link di riferimento di ExOv3:
https://learn.microsoft.com/it-it/powershell/module/exchange/?view=exchange-ps#powershell-v3-module

Una volta pronto per la distribuzione, sui branch unstable o master, inserirlo 
in una cartella "ExchangeOnLine" (da creare).
#>

<# *******************************************************************************
                                    HEADER
******************************************************************************* #>
# check execution policy
foreach ($item in (Get-ExecutionPolicy -List)) {
    if(($item.Scope -eq 'LocalMachine') -and ($item.ExecutionPolicy -cne 'Bypass')) {
        Write-Host "No enough privileges: open a PowerShell terminal with admin privileges and run the following cmdlet:`n"
        Write-Host -ForegroundColor Cyan "Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope LocalMachine -Force`n"
        Write-Host -NoNewline "Afterwards restart this script."
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
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# importing third party modules
$ErrorActionPreference= 'Stop'
do {
    try {
        Import-Module -Name "$workdir\Modules\Forms.psm1"
        Import-Module ExchangeOnlineManagement
    } catch {
        if (!(((Get-InstalledModule).Name) -contains 'ExchangeOnlineManagement')) {
            Install-Module ExchangeOnlineManagement -Confirm:$False -Force
            [System.Windows.MessageBox]::Show("Installed [ExchangeOnlineManagement] module: click Ok to restart the script",'RESTART','Ok','warning') > $null
            $ThirdParty = 'Ko'
        } else {
            [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error') > $null
            Write-Output "`nError: $($error[0].ToString())"
            Pause
            exit
        }
    }
} while ($ThirdParty -eq 'Ko')
$ErrorActionPreference= 'Inquire'


<# *******************************************************************************
                                    QUERYING
******************************************************************************* #>
Write-Host -NoNewline "Connecting to ExchangeOnLine... "
try {
    Connect-ExchangeOnline -ShowBanner:$false
    Write-Host -ForegroundColor Green "Ok"   
}
catch {
    Write-Host -ForegroundColor Red "Ko"
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    exit
}

Write-Host -NoNewline "Fetching mailbox list..."
$EXOlist = Get-EXOMailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize unlimited | ForEach-Object {
    Write-Host -NoNewline '.'        
    New-Object -TypeName PSObject -Property @{
        ID          = "$($_.Id)"
        TYPE        = "$($_.RecipientType)"
        OBJ_ID      = "$($_.ExternalDirectoryObjectId)"
        UPN         = "$($_.UserPrincipalName)"
        DISPLAY     = "$($_.DisplayName)"
    } | Select ID, TYPE, OBJ_ID, UPN, DISPLAY
}
Write-Host -ForegroundColor Green " Done"

Write-Host -NoNewline "Searching orphaned..."
$orphaned = Get-EXOMailbox -InactiveMailboxOnly -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize unlimited | ForEach-Object {
    Write-Host -NoNewline '.'        
    New-Object -TypeName PSObject -Property @{
        OBJ_ID      = "$($_.ExternalDirectoryObjectId)"
        UPN         = "$($_.UserPrincipalName)"
    } | Select OBJ_ID, UPN
}
Write-Host -ForegroundColor Green " Done"
if ($orphaned.Count -gt 0) {
    Clear-Host
    Write-Host -ForegroundColor Red "The following mailboxes result inactive"
    Pause
    Write-Host "`n"
    foreach ($item in $orphaned) {
        Write-Host -ForegroundColor Blue "[$($item.OBJ_ID)] $($item.UPN)"
    }
    Pause
}

Write-Host -NoNewline "Gathering mailbox details..."
$EXOdetailed = @{}
$totKeys = $EXOlist.Count
$counterKeys = 0
$parsebar = ProgressBar
foreach ($entity in $EXOlist) {
    $counterKeys ++
    Write-Host -NoNewline '.'
    $EXOdetailed["$entity.OBJ_ID"] = @{}


    # progressbar
    $percent = ($counterKeys / $totKeys)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Mailbox {0} out of {1} parsed [{2}%]" -f ($counterKeys, $totKeys, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}
Write-Host -ForegroundColor Green " Done"
$parsebar[0].Close()















Write-Host -NoNewline "Check sending perms..."
Get-EXORecipientPermission -ResultSize unlimited
Write-Host -ForegroundColor Green " Done"


Write-Host -NoNewline "Check ownership perms..."
$EXOroles = @{}
foreach ($Identity in $EXOlist.UPN) {
    Write-Host -NoNewline '.'
    $PermissionList = Get-EXOMailboxPermission -Identity $Identity
    if ($PermissionList.Count -gt 0) {
        <# Action to perform if the condition is true #>
    } else {
        $EXOroles["$Identity"] = @{ SELF = 'FullAccess' }
    }
}
Write-Host -ForegroundColor Green " Done"


<#
Get-EXOMailboxPermission -Identity it.fo@contoso.net

Identity        : It-Fornitori
User            : NT AUTHORITY\SELF
AccessRights    : {FullAccess, ExternalAccount, ReadPermission}
IsInherited     : False
Deny            : False
InheritanceType : Descendents

Identity        : It-Fornitori
User            : dario.corrada@contoso.net
AccessRights    : {FullAccess}
IsInherited     : False
Deny            : False
InheritanceType : All
#>

Disconnect-ExchangeOnline -Confirm:$false