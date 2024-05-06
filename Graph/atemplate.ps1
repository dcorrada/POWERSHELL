<#
This is a template for the header of those scripts that will use the Microsoft 
Graph module and wil access through the authenticator method (OTP and/or 
renewal of a token)
#>

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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AzureAD\\AssignedLicenses\.ps1$" > $null
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
    Import-Module -Name "$workdir\Modules\Forms.psm1"    
} catch {
    if (!(((Get-InstalledModule).Name) -contains 'Microsoft.Graph')) {
        Install-Module Microsoft.Graph -Scope AllUsers
        [System.Windows.MessageBox]::Show("Installed [MIcrosoft.Graph] module: please restart the script",'RESTART','Ok','warning')
        exit
    } else {
        [System.Windows.MessageBox]::Show("Error importing modules",'ABORTING','Ok','Error')
        Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
        Pause
        exit
    }
}
$ErrorActionPreference= 'Inquire'

Connect-MgGraph -Scopes 'User.Read.All'
<#
Alternativamente a questo cmdlet possibile accedere a Graph tramite 
un'app di autorizzazione Azure AD.

*** Creare un'applicazione nell'Azure Portal ***
Accedi al servizio Azure AD B2C e seleziona "Registrazioni per l'app":
https://portal.azure.com/#view/Microsoft_AAD_B2CAdmin/TenantManagementMenuBlade/~/overview

Crea una nuova registrazione dell'applicazione, limitando l'accesso solo ad 
"Account solo in questa directory dell'organizzazione (Tenant singolo)".

*** Memorizza le credenziali dell'applicazione ***
Dopo aver creato l'applicazione, prendi nota di "ID applicazione" (Client ID) e 
genera un segreto (Client Secret) cliccando sulla voce "Credenziali client".

*** Concedi autorizzazioni all'applicazione ***
Assicurati che l'applicazione abbia le autorizzazioni appropriate per accedere 
alle informazioni utente tramite Microsoft Graph (ie scope "User.Read.All"). 
Puoi farlo nelle impostazioni delle autorizzazioni dell'applicazione (click su 
voce "Autorizzazioni API").

*** Utilizza le credenziali dell'applicazione per l'autenticazione ***
Nel tuo script PowerShell, utilizza le credenziali dell'applicazione per 
ottenere un token di accesso tramite il flusso Client Credentials:

# Credenziali dell'applicazione
$clientId = "Your-Client-Id"
$clientSecret = "Your-Client-Secret"
$tenantId = "Your-Tenant-Id"
# Autenticazione e ottenimento del token di accesso
$token = Get-MgAccessToken -ClientId $clientId -ClientSecret $clientSecret -TenantId $tenantId -Scopes "https://graph.microsoft.com/.default"
# Imposta il token di accesso per l'utilizzo nelle richieste
Set-MgAccessToken -AccessToken $token.AccessToken
# Esegui le operazioni su Microsoft Graph
Get-MgUser -All
#>


Get-MgUser -All -Top 30
<#
Find AzureAD/MSOnline equivalent cmdlets for Graph on:
https://learn.microsoft.com/en-us/powershell/microsoftgraph/azuread-msoline-cmdlet-map?view=graph-powershell-1.0
#>
