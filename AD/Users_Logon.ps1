<#
Name......: Users_Logon.ps1
Version...: 21.02.1
Author....: Dario CORRADA

This script retrieve a list of all users which have logged on computers
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\AD\\Users_Logon\.ps1$" > $null
$workdir = $matches[1]

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

# import Active Directory module
if (! (get-Module ActiveDirectory)) { Import-Module ActiveDirectory }

# see https://sid-500.com/2018/02/28/powershell-get-all-logged-on-users-per-computer-ou-domain-get-userlogon/
function Get-UserLogon {
    [CmdletBinding()]
    param([Parameter ()][String]$Computer,[Parameter ()][String]$OU,[Parameter ()][Switch]$All)
     
    $ErrorActionPreference="SilentlyContinue"
    $result=@()
    
    If ($Computer) {
        Invoke-Command -ComputerName $Computer -ScriptBlock {quser} | Select-Object -Skip 1 | Foreach-Object {
            $b=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
            If ($b[2] -like 'Disc*') {
                $array= ([ordered]@{
                    'User' = $b[0]
                    'Computer' = $Computer
                    'Date' = $b[4]
                    'Time' = $b[5..6] -join ' '
                })
                $result+=New-Object -TypeName PSCustomObject -Property $array
            } else {
                $array= ([ordered]@{
                    'User' = $b[0]
                    'Computer' = $Computer
                    'Date' = $b[5]
                     'Time' = $b[6..7] -join ' '
                })
                $result+=New-Object -TypeName PSCustomObject -Property $array
            }
        }
    }
     
    If ($OU) { 
        $comp=Get-ADComputer -Filter * -SearchBase "$OU" -Properties operatingsystem
        $count=$comp.count
        If ($count -gt 20) {  
            Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer"
        }
        foreach ($u in $comp) {
            Invoke-Command -ComputerName $u.Name -ScriptBlock {quser} | Select-Object -Skip 1 | ForEach-Object {
                $a=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
                If ($a[2] -like '*Disc*') {
                    $array= ([ordered]@{
                        'User' = $a[0]
                        'Computer' = $u.Name
                        'Date' = $a[4]
                        'Time' = $a[5..6] -join ' '
                    })
                    $result+=New-Object -TypeName PSCustomObject -Property $array
                } else {
                    $array= ([ordered]@{
                        'User' = $a[0]
                        'Computer' = $u.Name
                        'Date' = $a[5]
                        'Time' = $a[6..7] -join ' '
                    })
                    $result+=New-Object -TypeName PSCustomObject -Property $array
                }
     
            }
        }
    }
     
    If ($All) {
        $comp=Get-ADComputer -Filter * -Properties operatingsystem
        $count=$comp.count
        If ($count -gt 20) {
            Write-Warning "Search $count computers. This may take some time ... About 4 seconds for each computer ..."
        }
        foreach ($u in $comp) {
            Invoke-Command -ComputerName $u.Name -ScriptBlock {quser} | Select-Object -Skip 1 | ForEach-Object {
                $a=$_.trim() -replace '\s+',' ' -replace '>','' -split '\s'
                If ($a[2] -like '*Disc*') {
                    $array= ([ordered]@{
                        'User' = $a[0]
                        'Computer' = $u.Name
                        'Date' = $a[4]
                        'Time' = $a[5..6] -join ' '
                    })
                    $result+=New-Object -TypeName PSCustomObject -Property $array
                } else {
                    $array= ([ordered]@{
                        'User' = $a[0]
                        'Computer' = $u.Name
                        'Date' = $a[5]
                        'Time' = $a[6..7] -join ' '
                    })
                    $result+=New-Object -TypeName PSCustomObject -Property $array
                }
            }
        }
    }
    Write-Output $result
}

# getting domain name
$output = nslookup ls
$output[0] -match "Server:\s+[a-zA-Z_\-0-9]+\.([a-zA-Z\-0-9\.]+)$" > $null
$dominio = $matches[1]

# OU dialog box
$form_modalita = FormBase -w 300 -h 230 -text "OU DESTINATION"
$consulenti  = RadioButton -form $form_modalita -checked $false -x 30 -y 20 -text "Client Consulenti"
$milano = RadioButton -form $form_modalita -checked $false -x 30 -y 50 -text "Client Milano"
$torino = RadioButton -form $form_modalita -checked $false -x 30 -y 80 -text "Client Torino"
OKButton -form $form_modalita -x 90 -y 150 -text "Ok"
$result = $form_modalita.ShowDialog()

# get distinguished name suffix
$dnsuffix = ''
foreach ($dctag in $dominio.Split('.')) {
    $dnsuffix += ',DC=' + $dctag
}

# in "elseif" blocks modify $outarget prefix according to yours OU paths
if ($result -eq "OK") {
    if ($consulenti.Checked) {
        $outarget = 'OU=Client Consulenti,OU=Delegate' + $dnsuffix
    } elseif ($milano.Checked) {
        $outarget = 'OU=Client Milano,OU=Delegate' + $dnsuffix
    } elseif ($torino.Checked) {
        $outarget = 'OU=Client Torino,OU=Delegate' + $dnsuffix
    }    
}

Write-Host "OU.......: " -NoNewline
Write-Host $outarget -ForegroundColor Cyan

# retrieve logon list
Get-UserLogon -OU $outarget

Pause