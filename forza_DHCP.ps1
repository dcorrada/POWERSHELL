<#
Name......: forza_DHCP.ps1
Version...: 20.1.1
Author....: Dario CORRADA

Questo script serve per riconfigurare i settaggi di rete su DHCP
#>

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'

# configuro il DHCP
$dhcp_check = Get-NetIPInterface
foreach ($item in $dhcp_check) {
    if (($item.InterfaceAlias -eq 'Ethernet') -and ($item.AddressFamily -eq 'IPv4')) {
        if ($item.Dhcp -eq 'Disabled') {
            Write-Host "Abilito il DHCP"
            Set-NetIPInterface -InterfaceAlias 'Ethernet' -Dhcp Enabled -Confirm:$false
        }
    }
}

# Configuro il DNS
Set-DnsClientServerAddress -InterfaceAlias 'Ethernet' -ResetServerAddresses 
Set-DnsClientGlobalSetting -SuffixSearchList ''   

# ripulisco la cache dal DNS
Clear-DnsClientCache
