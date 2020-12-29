$CM12MP='MP.contoso.com'

$CMSiteCode='SR1'

$ErrorActionPreference = "SilentlyContinue" 

try 
{ 
#Get ccm cache path for later cleanup... 
    try 
    { 
        $ccmcache = ([wmi]"ROOT\ccm\SoftMgmtAgent:CacheConfig.ConfigKey='Cache'").Location 
    } catch {} 

#download ccmsetup.exe from MP 
    $webclient = New-Object System.Net.WebClient 
    $url = "http://$($CM12MP)/CCM_Client/ccmsetup.exe" 
    $file = "c:\windows\temp\ccmsetup.exe" 
    $webclient.DownloadFile($url,$file) 

#download CU2 Patch from MP , must be copied manually to the Client directory.
#$url = "http://$($CM12MP)/CCM_Client/configmgr2012ac-r2-kb2970177-x64.msp" 
#$file = "c:\windows\temp\configmgr2012ac-r2-kb2970177-x64.msp" 
#$webclient.DownloadFile($url,$file)

#stop the old sms agent service 
    stop-service 'ccmexec' -ErrorAction SilentlyContinue 

#Cleanup cache 
    if($ccmcache -ne $null) 
    { 
        try 
        { 
        dir $ccmcache '*' -directory | % { [io.directory]::delete($_.fullname, $true)  } -ErrorAction SilentlyContinue 
        } catch {} 
    } 

#Cleanup Execution History 
    Remove-Item -Path 'HKLM:\SOFTWARE\Wow6432Node\Microsoft\SMS\Mobile Client\*' -Recurse -ErrorAction SilentlyContinue 

#Cleanup App-V 4.6 Packages 
    try 
    { 
        (get-wmiobject -query "SELECT * FROM Package WHERE SftPath like '%' AND InUse = 'FALSE' " -namespace "root\Microsoft\appvirt\client") | % { start-process -wait sftmime.exe -argumentlist "delete package:$([char]34)$($_.Name)$([char]34) /global" }         
    } catch {} 

#kill existing instances of ccmsetup.exe 
    $ccm = (Get-Process 'ccmsetup' -ErrorAction SilentlyContinue) 
    if($ccm -ne $null) 
    { 
            $ccm.kill(); 
    } 

#run ccmsetup
if(test-path "c:\windows\temp\configmgr2012ac-r2-kb2970177-x64.msp") {
    $proc = Start-Process -FilePath 'c:\windows\temp\ccmsetup.exe' -PassThru -ArgumentList "Patch=""c:\windows\temp\configmgr2012ac-r2-kb2970177-x64.msp"" /mp:$($CM12MP) /source:http://$($CM12MP)/CCM_Client CCMHTTPPORT=80 RESETKEYINFORMATION=TRUE SMSSITECODE=$($CMSiteCode) SMSSLP=$($CM12MP) FSP=$($CM12MP)"
} else { 
    $proc = Start-Process -FilePath 'c:\windows\temp\ccmsetup.exe' -PassThru -ArgumentList "/mp:$($CM12MP) /source:http://$($CM12MP)/CCM_Client CCMHTTPPORT=80 RESETKEYINFORMATION=TRUE SMSSITECODE=$($CMSiteCode) SMSSLP=$($CM12MP) FSP=$($CM12MP)"
}  
    Sleep(5)
	"ccmsetup started..." 
} 

catch 
{ 
        "an Error occured..." 
        $error[0] 
} 