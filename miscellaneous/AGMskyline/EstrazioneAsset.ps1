# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\EstrazioneAsset\.ps1$" > $null
$workdir = $matches[1]
<# alternative for testing
$workdir = Get-Location
$workdir = $workdir.Path
#>

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Forms.psm1"

# accessing info
$uri_prefix = 'http://192.168.2.184/' # IP Snipe webserver
$token_file = $env:LOCALAPPDATA + '\SnipeIT.token'
$token_string = Get-Content $token_file

# header for request
$headers = @{
    'Authorization' = "Bearer $token_string"
    'Accept' = 'application/json'
    'Content-Type' = 'application/json'
}

# query parameters
$query_params = @(
    'limit=1000'
    'offset=0'
)
$uri_suffix = '?' + ($query_params -join '&')
$uri = $uri_prefix + 'api/v1/hardware' + $uri_suffix
$ErrorActionPreference= 'Stop'
Try {
    $rawdata = Invoke-RestMethod -Uri $uri -Method Get -Headers $headers
    $ErrorActionPreference= 'Inquire'
}
Catch {
    [System.Windows.MessageBox]::Show("$($error[0].ToString())",'ABORTING','Ok','Error')
    exit
}

# get status list
Write-Host -NoNewline "Looking for status... "
$status_list = @()
foreach ($record in $rawdata.rows) {
    $status = $record.status_label.name
    if (!($status_list -contains $status)) {
        $status_list += $status
    }
}
Write-Host -ForegroundColor Green "DONE"
$adialog = FormBase -w 275 -h ((($status_list.Count-1) * 30) + 150) -text "SELECT STATUS"
$they = 20
$choices = @()
foreach ($item in $status_list) {
    if (('Assegnato', 'da Assegnare', 'in transito') -contains $item) {
        $choices += CheckBox -form $adialog -checked $true -x 20 -y $they -text $item
    } else {
        $choices += CheckBox -form $adialog -checked $false -x 20 -y $they -text $item
    }
    
    $they += 30 
}
OKButton -form $adialog -x 75 -y ($they + 20) -text "Ok" | Out-Null
$result = $adialog.ShowDialog()
$selected_status = @()
foreach ($item in $choices) {
    if ($item.Checked) {
        $selected_status += $item.Text
    }
}

# collecting data
$enc = [System.Text.Encoding]::UTF8
$fetched_data = @()
$i = 1
$totrec = $rawdata.rows.Count
$parsebar = ProgressBar
foreach ($record in $rawdata.rows) {
    $fetched_record = @{}
    $fetched_record.STATUS = $record.status_label.name
    if ($selected_status -contains $fetched_record.STATUS) {
        $fetched_record.MODEL = $record.manufacturer.name + ' ' + $record.model.name
        $fetched_record.NAME = $record.name
        $fetched_record.SERIAL = $record.serial
        if ($record.order_number) {
            $fetched_record.NUMPROT = $record.order_number
        } else {
            $fetched_record.NUMPROT = 'na'
        }
        $fetched_record.LOCATION = $record.location.name
        $fetched_record.UPTODATE = $record.updated_at.datetime | Get-Date -format "yyyy-MM-dd"
        if ($record.assigned_to.name) {
            $fetched_record.ASSIGNED_TO = $enc.GetString($enc.GetBytes($record.assigned_to.name))
            $fetched_record.ASSIGNED_TO_USR = $record.assigned_to.username
        } else {
            $fetched_record.ASSIGNED_TO = 'NULL'
            $fetched_record.ASSIGNED_TO_USR = 'NULL'
        }
        if ($record.custom_fields.'CPU generazione'.value) {
            $fetched_record.CPU = $record.custom_fields.Processore.value + '_' + $record.custom_fields.'CPU generazione'.value
        } else {
            $fetched_record.CPU = $record.custom_fields.Processore.value
        }
        $fetched_record.RAM = $record.custom_fields.'RAM (GB)'.value
        $fetched_record.SSD = $record.custom_fields.'Disco (GB)'.value

        $fetched_data += $fetched_record
    }
    $i++
    # progress
    $percent = (($i-1) / $totrec)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Collecting {0} out of {1} records [{2}%]" -f (($i-1), $totrec, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()    
}
$parsebar[0].Close()

# writing output file
Write-Host -NoNewline "Writing output file... "
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-EstrazioneAsset.csv'


$i = 1
$totrec = $fetched_data.Count
$parsebar = ProgressBar
foreach ($item in $fetched_data) {
    $string = ("AGM{0:d5};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12}" -f ($i,$item.NAME,$item.STATUS,$item.MODEL,$item.SERIAL,$item.NUMPROT,$item.LOCATION,$item.UPTODATE,$item.ASSIGNED_TO,$item.ASSIGNED_TO_USR,$item.CPU,$item.RAM,$item.SSD))
    $string = $string -replace ';_;', ';NULL;'
    $string = $string -replace ';\s*;', ';NULL;'
    $string = $string -replace ';\s*$', ';NULL'
    $string = $string -replace '&#039;', "'"
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

