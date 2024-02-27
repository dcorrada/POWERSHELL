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
    Import-Module FOOBARBAZ
} catch {
    if (!(((Get-InstalledModule).Name) -contains 'FOOBARBAZ')) {
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

<#
Find AzureAD/MSOnline equivalent cmdlets for Graph on:

https://learn.microsoft.com/en-us/powershell/microsoftgraph/azuread-msoline-cmdlet-map?view=graph-powershell-1.0

E.g.: how to extract all user list (aka old Get-AzureADUser)
Connect-MgGraph -Scopes 'User.Read.All'
Get-MgUser -ConsistencyLevel eventual -Count userCount -Filter "startsWith(DisplayName, 'a')" -Top 1
#>