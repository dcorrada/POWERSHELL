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

# percorsi di rete
$target_paths = @{
    'ASSEGNAZIONI' = '\\192.168.2.251\Share\AREA HR\Selezione\ASSEGNAZIONI - VISITE - SICUREZZA\VARIE\Strumenti aziendali\Assegnazioni dipendenti\Assegnazione CELL'
    'CONSEGNE' = '\\192.168.2.251\Share\AREA HR\Selezione\ASSEGNAZIONI - VISITE - SICUREZZA\VARIE\Strumenti aziendali\Consegne  - resi dipendenti\Cellulari'
}

$aword = New-Object -ComObject Word.application
$parseddata = @()
foreach ($status in $target_paths.Keys) {
    # recupero la lista delle schede
    $filelist = Get-ChildItem -Path ($target_paths[$status] + '\*.*')

    Clear-Host
    Write-Host -ForegroundColor Cyan "*** ANALISI SCHEDE $status ***"
    Start-Sleep -Seconds 2

    $tot = $filelist.Count
    $i = 0
    $parsebar = ProgressBar
    foreach ($item in $filelist) {
        $filename = $item.Name
        $timestamp = $item.LastWriteTime | Get-Date -f yyyy-MM-dd
    
        $gotcha = $filename -imatch "telefon"
        if ($gotcha -eq $true) {
            $arecord = @{
                'FILENAME' = $filename
                'UPDATED' = $timestamp
                'STATUS' =  'NULL'
                'FULLNAME' =  'NULL'
                'DESCRIPTION' =  'NULL'
            }

            if ($filename -imatch "assegnazione") {
                $arecord.STATUS = 'ASSEGNATO'
            } elseif ($filename -imatch "consegna") {
                $arecord.STATUS = 'CONSEGNATO'
            }

            $ext = $filename -match "\.docx$"
            if ($ext -eq $true) {
                Write-Host -ForegroundColor Blue "Parsing [$filename]..."

                $ErrorActionPreference= 'Stop'
                try {
                    $infile = "C:\Users\$env:USERNAME\Downloads\$filename"
                    Copy-Item -Path $item.FullName -Destination $infile
                    $adoc = $aword.Documents.Open($infile)

                    for ($par = 1; $par -lt $adoc.Paragraphs.Count; $par++) {
                        $string = $adoc.Paragraphs[$par].range.Text
                        $gotcha = $string -match "^Egr\. (Sig\.|Sig\.ra|Dott\.)\s+([a-zA-Z\s]+),*\s{2,}"
                        if ($gotcha -eq $true) {
                            $fullname = $matches[2] -replace '\s+$', ''
                            $fullname = $fullname -replace '^ra\s+', ''
                            $arecord.FULLNAME = $fullname
                        }
                        $gotcha = $string -imatch "^Tipologia materiale: (.+)$"
                        if ($gotcha -eq $true) {
                            $device = $matches[1] -replace '\s+$', ''
                            $arecord.DESCRIPTION = $device
                        }
                    }

                    $adoc.close()
                    Remove-Item -Path $infile
                } catch {
                    $amsg = "Error: $($error[0].ToString())"
                    $answ = [System.Windows.MessageBox]::Show("$amsg",'ABORTING?','YesNo','Error')
                    if ($answ -eq 'Yes') {
                        Exit
                    }
                }
                $ErrorActionPreference= 'Inquire'
            } else {                
                Write-Host -ForegroundColor Yellow "Skipping [$filename]..."
                # non sono file DOCX, vedere come parsare i PDF...
            }

            $parseddata += $arecord
        }
        
        # progress
        $i++
        $percent = ($i / $tot)*100
        if ($percent -gt 100) {
            $percent = 100
        }
        $formattato = '{0:0.0}' -f $percent
        [int32]$progress = $percent   
        $parsebar[2].Text = ("Record {0} out of {1} parsed [{2}%]" -f ($i, $tot, $formattato))
        if ($progress -ge 100) {
            $parsebar[1].Value = 100
        } else {
            $parsebar[1].Value = $progress
        }
        [System.Windows.Forms.Application]::DoEvents()
    }
    $parsebar[0].Close()
}
$aword.Quit()

# applico patch sui record non parsati correttamento
Clear-Host
Write-Host -ForegroundColor Cyan "*** AGGIORNAMENTO RECORD ***"
Start-Sleep -Seconds 2

[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Open Patch File"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
$OpenFileDialog.filename = 'SchedeTelefoniPATCHED'
$OpenFileDialog.ShowDialog() | Out-Null
$infile = $OpenFileDialog.filename

$ErrorActionPreference= 'Stop'
try {
    Import-Module ImportExcel
} catch {
    Install-Module ImportExcel -Confirm:$False -Force
    Import-Module ImportExcel
}
$ErrorActionPreference= 'Inquire'
$rawpatches = Import-Excel -Path $infile -WorksheetName 'SchedeTelefoniPATCHED'
$parsedpatches = @{}
foreach ($itemnew in $rawpatches) {
    $parsedpatches[$itemnew.FILENAME] = @{
        UPDATED = $itemnew.UPDATED
        STATUS = $itemnew.STATUS
        FULLNAME = $itemnew.FULLNAME
        DESCRIPTION = $itemnew.DESCRIPTION
    }
}

$updata = @()
foreach ($candidate in $parseddata) {
    $akey = $candidate.FILENAME

    Write-Host -NoNewline "Check [$akey]... "
    if ($parsedpatches.ContainsKey($akey)) {
        $arecord = @{
            'FILENAME' = $akey
            'UPDATED' = $parsedpatches[$akey].UPDATED | Get-Date -f yyyy-MM-dd
            'STATUS' =  $parsedpatches[$akey].STATUS
            'FULLNAME' =  $parsedpatches[$akey].FULLNAME
            'DESCRIPTION' =  $parsedpatches[$akey].DESCRIPTION
        }
        Write-Host -ForegroundColor Cyan 'UPDATED'
    } else {
        $arecord = $candidate
        Write-Host -ForegroundColor Green 'DONE'
    }

    $updata += $arecord
    Start-Sleep -Milliseconds 100
}


# writing output file
Write-Host -NoNewline "Writing output file... "
$outfile = "C:\Users\$env:USERNAME\Downloads\" + (Get-Date -format "yyMMdd") + '-SchedeTelefoni.csv'

$i = 1
$totrec = $updata.Count
$parsebar = ProgressBar
foreach ($item in $updata) {
    $string = ("AGM{0:d5};{1};{2};{3};{4};{5}" -f ($i,$item.FILENAME,$item.UPDATED,$item.STATUS,$item.FULLNAME,$item.DESCRIPTION))
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
