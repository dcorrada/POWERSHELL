<#
This is a template for testing those scripts that will use the Microsoft Graph 
module and wil access through the authenticator method
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Graph\\PowerShell_SDK\\atemplate\.ps1$" > $null
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

<# 
*******************************************************************************
                            CREDENTIALS MANAGEMENT
*******************************************************************************
This section retrieve all informations needed to connect using client secret 
credential method as described in:
https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph?view=graph-powershell-1.0#example-8-using-client-secret-credentials

Alternatively you can connect interactively via delegate access as described in:
https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/connect-mggraph?view=graph-powershell-1.0#example-4-delegated-access-custom-application-for-microsoft-graph-powershell

Registration of an ad hoc Azure app is required for such athentication methods.
For more details see AZURE_APP.txt file in the parent folder.
#>
Write-Host -NoNewline "Credential management... "
$pswout = PowerShell.exe -file "$workdir\Graph\AppKeyring.ps1"
if ($pswout.Count -eq 4) {
    $UPN = $pswout[0]
    $clientID = $pswout[1]
    $tenantID = $pswout[2]
    $Clientsecret = $pswout[3]
    Write-Host -ForegroundColor Green 'Ok'
} else {
    [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "Ko"
    Pause
    exit
}
$ClientSecretCredential = New-Object System.Management.Automation.PSCredential($clientID, (ConvertTo-SecureString $Clientsecret -AsPlainText -Force))

# connect using client secret (not tested yet)
# Connect-MgGraph -TenantId $tenantID -ClientSecretCredential $ClientSecretCredential

# connect interactively via delegate access
Connect-MgGraph -ClientId $clientID -TenantId $tenantID 



<# *******************************************************************************
                                    MAIN
******************************************************************************* #>

Get-MgUser -All -Top 10
<#
Find AzureAD/MSOnline equivalent cmdlets for Graph on:
https://learn.microsoft.com/en-us/powershell/microsoftgraph/azuread-msoline-cmdlet-map?view=graph-powershell-1.0
#>


$infoLogout = Disconnect-Graph
<#
Once you're signed in, you'll remain signed in until you invoke Disconnect-MgGraph. 
Microsoft Graph PowerShell automatically refreshes the access token for you and 
sign-in persists across PowerShell sessions because Microsoft Graph PowerShell 
securely caches the token.

Use Disconnect-MgGraph cmdlet to sign out.
#>