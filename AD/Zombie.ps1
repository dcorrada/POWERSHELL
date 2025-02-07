<#
Name......: Zombie.ps1
Version...: 22.02.1
Author....: Dario CORRADA

This script looks for those AD account that need to be dismissed (ie don't have a Microsoft account) and save a list in a CSV file
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AD\\Zombie\.ps1$" > $null
$workdir = $matches[1]

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# check Active Directory module
if ((Get-Module -Name ActiveDirectory -ListAvailable) -eq $null) {
    $ErrorActionPreference= 'Stop'
    try {
        Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability –Online
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to install RSAT"
        Pause
        Exit
    }
    $ErrorActionPreference= 'Inquire'
}

# import the AzureAD module
$ErrorActionPreference= 'Stop'
try {
    Import-Module MSOnline
} catch {
    Install-Module MSOnline -Confirm:$False -Force
    Import-Module MSOnline
}
$ErrorActionPreference= 'Inquire'

# retrieve the list of available OUs
$ou_available = Get-ADOrganizationalUnit -Filter *
$hsize = 150 + (25 * $ou_available.Count)
$form_panel = FormBase -w 300 -h $hsize -text "AVAILABLE OU"
Label -form $form_panel -x 10 -y 20 -text 'Select OU for looking for:' | Out-Null 
$vpos = 45
$boxes = @()
foreach ($elem in ($ou_available.Name | Sort-Object)) {
    $boxes += CheckBox -form $form_panel -checked $false -x 20 -y $vpos -text $elem
    $vpos += 25
}
$vpos += 20
OKButton -form $form_panel -x 90 -y $vpos -text "Ok" | Out-Null
$result = $form_panel.ShowDialog()

# retrieve users list from selected OUs
$rawdata = @{}
foreach ($box in $boxes) {
    if ($box.Checked -eq $true) {
        $OUname = $box.Text
        Write-Host -NoNewline "Retrieving userlist from [$OUname]..."
        for ($i = 0; $i -lt $ou_available.Count; $i++) {
            if ($OUname -eq $ou_available[$i].Name) {
                $OUpath = $ou_available[$i].DistinguishedName
                $user_list = Get-ADUser -Filter * -SearchBase $OUpath | Select-object Name,UserPrincipalName
                for ($k = 0; $k -lt $user_list.Count; $k++) {
                    $user_list[$k].UserPrincipalName -match "^(.+)@.+$" > $null
                    $akey = $matches[1]
                    $rawdata.$akey = @{  
                        FullName = $user_list[$k].Name      
                        OU = $OUname
                        UsrName = $akey
                    }
                    Write-Host -NoNewline "."
                }
                Write-Host -ForegroundColor Green " OK"
            }
        }
    }
}

# looking for any Azure AD account
Connect-MsolService
$outfile = "$env:USERPROFILE\Downloads\ZOMBIE.csv"
"FULLNAME;USERNAME;OU" | Out-File $outfile -Encoding utf8
foreach ($item in ($rawdata.Keys | Sort-Object)) {
    Write-Host -NoNewline "Checking [$item]..."
    $ErrorActionPreference= 'Stop'
    try {
        # re-edit your current domain suffix before running this script
        Get-MsolUser -UserPrincipalName "$item@foobarbaz.net" > $null
        Write-Host -ForegroundColor Green " OK"
    }
    catch {
        Write-Host -ForegroundColor Red " KO"
        $new_record = @(
            $rawdata.$item.FullName,
            $rawdata.$item.UsrName,
            $rawdata.$item.OU
        )
        $new_string = [system.String]::Join(";", $new_record)
        $new_string | Out-File $outfile -Encoding utf8 -Append
    }
    $ErrorActionPreference= 'Inquire'
}
Pause
