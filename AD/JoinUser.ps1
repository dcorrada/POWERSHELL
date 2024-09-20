<#
Name......: JoinUser.ps1
Version...: 21.06.1
Author....: Dario CORRADA

This script grants local admin privileges to an existing domain user
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
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"


<# *******************************************************************************
                                    BODY
******************************************************************************* #>

$answ = [System.Windows.MessageBox]::Show("Are you looking for an AD user?",'ACCOUNT','YesNo','Info')
if ($answ -eq "Yes") {

    # dialog form
    $form = FormBase -w 350 -h 175 -text "ACCOUNT"
    Label -form $form -x 10 -y 20 -w 90 -text 'Username:'  | Out-Null
    $usrname = TxtBox -form $form -x 100 -y 20 -w 200
    Label -form $form -x 10 -y 50 -w 90 -text 'Password:'  | Out-Null
    $passwd = TxtBox -form $form -x 100 -y 50 -w 200 -masked $true
    OKButton -form $form -x 120 -y 90 -text "Ok" | Out-Null
    $result = $form.ShowDialog()

    # add domain prefix to username
    $username = $usrname.Text
    $thiscomputer = Get-WmiObject -Class Win32_ComputerSystem
    $fullname = $thiscomputer.Domain + '\' + $username

    # test user
    [reflection.assembly]::LoadWithPartialName("System.DirectoryServices.AccountManagement") > $null
    $principalContext = [System.DirectoryServices.AccountManagement.PrincipalContext]::new([System.DirectoryServices.AccountManagement.ContextType]'Machine',$env:COMPUTERNAME)
    if ($principalContext.ValidateCredentials($fullname,$passwd.Text)) {
        Write-Host -ForegroundColor Green "User OK"
        
        # granting local admin privileges
        try {
            Add-LocalGroupMember -Group "Administrators" -Member $fullname
        }
        catch {
            [System.Windows.MessageBox]::Show("Cannot granting admin privilege to $username",'ACCOUNT','Ok','Error')
        }
    } else {
        [System.Windows.MessageBox]::Show("Invalid credentials for $username",'ACCOUNT','Ok','Error')
    }
}