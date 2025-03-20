<#
https://learn.microsoft.com/en-us/answers/questions/748166/ms-graph-get-owned-distribution-lists-for-a-user

I gruppi sembrano essere ordinati in base a questa categoria di attributi su Graph

                            groupType       mailEnabled     securityEnabled
------------------------------------------------------------------------------
Microsoft 365               unified         true            false
Security                    null            false           true
Mail-enabled security       null            true            true
Distribution List           null            true            false
#>


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
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent | Split-Path -Parent

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"
Import-Module ExchangeOnlineManagement

# retrieve credentials
Write-Host -NoNewline "Credential management... "
$pswout = PowerShell.exe -file "$workdir\Graph\AppKeyring.ps1"
if ($pswout.Count -eq 4) {
    $UPN = $pswout[0]
    $clientID = $pswout[1]
    $tenantID = $pswout[2]
    Write-Host -ForegroundColor Green ' Ok'
} else {
    [System.Windows.MessageBox]::Show("Error connecting to PSWallet",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

# connect to Tenant
Write-Host -NoNewline "Connecting to the Tenant..."
$ErrorActionPreference= 'Stop'
Try {
    $splash = Connect-MgGraph -ClientId $clientID -TenantId $tenantID 
    Write-Host -ForegroundColor Green ' Ok'
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("Error connecting to the Tenant",'ABORTING','Ok','Error')
    Write-Host -ForegroundColor Red "ERROR: $($error[0].ToString())"
    Pause
    exit
}

# get available groups and related members
Write-Host -NoNewline "Retrieving group list... "
$GroupList = Get-MgGroup -All  -Property Id, DisplayName, Description, Mail, CreatedDateTime, GroupTypes, mailEnabled, securityEnabled `
    | Select-Object Id, DisplayName, Description, Mail, CreatedDateTime, GroupTypes, mailEnabled, securityEnabled
Write-Host -ForegroundColor Green 'Ok'

$fetched_data = @()
Write-Host "Looking for members..."
foreach ($currentGroup in $GroupList) {
    Write-Host -NoNewline -ForegroundColor Blue "   $($currentGroup.DisplayName)"
    $MemberShip = Get-MgGroupMember -All -GroupID $currentGroup.Id  | Select-Object -Property Id, AdditionalProperties
    if ($MemberShip.Count -gt 0) {
        foreach ($currentMember in $MemberShip) {
            Write-Host -NoNewline '.'
            if ($currentMember.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user') {
                $fetched_record = @{
                    DLNAME          = $currentGroup.DisplayName
                    DLDESC          = $currentGroup.Description -replace '"', ' '
                    DLMAIL          = $currentGroup.Mail
                    DLCREATED       = $currentGroup.CreatedDateTime | Get-Date -format "yyyy-MM-dd"
                    DLSECURITY      = $currentGroup.securityEnabled
                    DLMAILENABLED   = $currentGroup.mailEnabled
                    DLTYPE          = $currentGroup.GroupTypes -join '.'
                    FULLNAME        = $currentMember.AdditionalProperties.displayName
                    EMAIL           = $currentMember.AdditionalProperties.userPrincipalName
                }
                $fetched_data += $fetched_record
            } elseif ($currentMember.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.group') {
                $fetched_record = @{
                    DLNAME          = $currentGroup.DisplayName
                    DLDESC          = $currentGroup.Description -replace '"', ' '
                    DLMAIL          = $currentGroup.Mail
                    DLCREATED       = $currentGroup.CreatedDateTime | Get-Date -format "yyyy-MM-dd"
                    DLSECURITY      = $currentGroup.securityEnabled
                    DLMAILENABLED   = $currentGroup.mailEnabled
                    DLTYPE          = $currentGroup.GroupTypes -join '.'
                    FULLNAME        = $currentMember.AdditionalProperties.displayName
                    EMAIL           = $currentMember.AdditionalProperties.mail
                }
                $fetched_data += $fetched_record
            } elseif ($currentMember.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.orgContact') {
                $aContact = Get-MgContact -OrgContactId $currentMember.Id
                $fetched_record = @{
                    DLNAME          = $currentGroup.DisplayName
                    DLDESC          = $currentGroup.Description -replace '"', ' '
                    DLMAIL          = $currentGroup.Mail
                    DLCREATED       = $currentGroup.CreatedDateTime | Get-Date -format "yyyy-MM-dd"
                    DLSECURITY      = $currentGroup.securityEnabled
                    DLMAILENABLED   = $currentGroup.mailEnabled
                    DLTYPE          = $currentGroup.GroupTypes -join '.'
                    FULLNAME        = $aContact.DisplayName
                    EMAIL           = $aContact.Mail
                }
                $fetched_data += $fetched_record
            } else {
                Write-Host -ForegroundColor Yellow @"

    UNEXPECTED DATA...: $($currentMember.AdditionalProperties.'@odata.type')
    PROPERTIES........: $($currentMember.AdditionalProperties.Keys -join '|')

"@
                Pause
            } 
        }
    } else {
        $fetched_record = @{
            DLNAME          = $currentGroup.DisplayName
            DLDESC          = $currentGroup.Description -replace '"', ' '
            DLMAIL          = $currentGroup.Mail
            DLCREATED       = $currentGroup.CreatedDateTime | Get-Date -format "yyyy-MM-dd"
            DLSECURITY      = $currentGroup.securityEnabled
            DLMAILENABLED   = $currentGroup.mailEnabled
            DLTYPE          = $currentGroup.GroupTypes -join '.'
            FULLNAME        = 'nobody'
            EMAIL           = 'none'
        }
        $fetched_data += $fetched_record
    }
    Write-Host -ForegroundColor Green ' Ok'

    Disconnect-ExchangeOnline -Confirm:$false
}

# disconnect from Tenant
$infoLogout = Disconnect-Graph

# Dynamic Distribution Lists
Write-Host -NoNewline "`nFetching dynamic DLs..."
try {
    Connect-ExchangeOnline -ShowBanner:$false
    $eolok = $true
    Write-Host -ForegroundColor Green "LOGGED"
}
catch {
    Write-Host -ForegroundColor Yellow " SKIPPED`n(unable to connect to ExchangeOnLine)"
    $eolok = $false
    Pause
}
if ($eolok) {
    foreach ($dyndl in Get-DynamicDistributionGroup) {
        Write-Host -NoNewline -ForegroundColor Blue "   $($dyndl.DisplayName)"
        foreach ($member in Get-DynamicDistributionGroupMember -Identity $dyndl.DisplayName) {
            $fetched_record = @{
                DLNAME          = $dyndl.DisplayName
                DLDESC          = $dyndl.Notes
                DLMAIL          = $dyndl.PrimarySmtpAddress
                DLCREATED       = $dyndl.WhenCreated | Get-Date -format "yyyy-MM-dd"
                DLSECURITY      = 'False'
                DLMAILENABLED   = 'True'
                DLTYPE          = 'Dynamic'
                FULLNAME        = $member.DisplayName
                EMAIL           = $member.PrimarySmtpAddress
            }
            $fetched_data += $fetched_record
            Write-Host -NoNewline '.'
        }
        Write-Host -ForegroundColor Green ' Ok'
    }
}


# writing output file
Write-Host -NoNewline "Writing output file... "
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-DLmembers.csv'

$i = 1
$totrec = $fetched_data.Count
$parsebar = ProgressBar
foreach ($item in $fetched_data) {
    $string = ("AGM{0:d5};{1};{2};{3};{4};{5};{6};{7};{8};{9}" -f ($i,$item.DLNAME,$item.DLDESC,$item.DLMAIL,$item.DLCREATED,$item.DLSECURITY,$item.DLMAILENABLED,$item.DLTYPE,$item.FULLNAME,$item.EMAIL))
    $string = $string -replace ';\s*;\s*;', ';NULL;NULL;'
    $string = $string -replace ';\s*;', ';NULL;'
    $string = $string -replace ';+\s*$', ';NULL'
    $string = $string -replace ';', '";"'
    $string = '"' + $string + '"'
    $string = $string -replace '"NULL"', 'NULL'
    $string | Out-File $outfile -Encoding utf8 -Append
    $i++

    # progress
    $percent = (($i-1) / $totrec)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Writing {0} out of {1} records [{2}%]" -f (($i-1), $totrec, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}
$parsebar[0].Close()
Write-Host -ForegroundColor Green "DONE"
