<#
Name......: Force_DHCP.ps1
Version...: 20.1.1
Author....: Dario CORRADA

This script reconfigure network settings to DHCP
#>

# setting DHCP profile
$dhcp_check = Get-NetIPInterface
foreach ($item in $dhcp_check) {
    if (($item.InterfaceAlias -eq 'Ethernet') -and ($item.AddressFamily -eq 'IPv4')) {
        if ($item.Dhcp -eq 'Disabled') {
            Write-Host "Enabling DHCP"
            Set-NetIPInterface -InterfaceAlias 'Ethernet' -Dhcp Enabled -Confirm:$false
        }
    }
}

# setting DNS
Set-DnsClientServerAddress -InterfaceAlias 'Ethernet' -ResetServerAddresses 
Set-DnsClientGlobalSetting -SuffixSearchList ''   

# clear DNS cache
Clear-DnsClientCache
