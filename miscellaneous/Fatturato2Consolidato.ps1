<# 
NAME......: Fatturato2Consolidato.ps1
VERSION...: 22.07.1
AUTHOR....: Dario CORRADA
#>

<# chunck per lanciare lo script con privilegi elevati
# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}
#>

# funzione per disegnare una progress bar
function ProgressBar {
    $form_bar = New-Object System.Windows.Forms.Form
    $form_bar.Text = "PROGRESS"
    $form_bar.Size = New-Object System.Drawing.Size(600,200)
    $form_bar.StartPosition = 'CenterScreen'
    $form_bar.Topmost = $true
    $form_bar.MinimizeBox = $false
    $form_bar.MaximizeBox = $false
    $form_bar.FormBorderStyle = 'FixedSingle'
    $font = New-Object System.Drawing.Font("Arial", 12)
    $form_bar.Font = $font
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(20,20)
    $label.Size = New-Object System.Drawing.Size(550,30)
    $form_bar.Controls.Add($label)
    $bar = New-Object System.Windows.Forms.ProgressBar
    $bar.Style="Continuous"
    $bar.Location = New-Object System.Drawing.Point(20,70)
    $bar.Maximum = 101
    $bar.Size = New-Object System.Drawing.Size(550,30)
    $form_bar.Controls.Add($bar)
    $form_bar.Show() | out-null
    return @($form_bar, $bar, $label)
}

# importo roba grafica
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# finestra di dialogo per caricare il file Excel del fatturato prodotto da Diamante
[System.Reflection.Assembly]::LoadWithPartialName('System.windows.forms') | Out-Null
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Carica Fatturato"
$OpenFileDialog.initialDirectory = "C:\Users\$env:USERNAME\Desktop"
$OpenFileDialog.filter = 'Excel file (*.xlsx)| *.xlsx'
$OpenFileDialog.ShowDialog() | Out-Null
$rawfile = $OpenFileDialog.filename

# importo il file Excel del fatturato
$ErrorActionPreference= 'Stop'
try {
    Import-Module ImportExcel
} catch {
    Install-Module ImportExcel -Confirm:$False -Force
    Import-Module ImportExcel
}
$ErrorActionPreference= 'Inquire'
$rawdata = Import-Excel $rawfile 

# verifico che l'header (aka il nome delle singole colonne) sia conforme
$campi = $rawdata | Get-Member
foreach ($currentItemName in ('Tipo', 'Numero Doc', 'Del', 'Cod', 'Cliente', 'Cod Art', 'Articolo', 'Conto', 'Qta', 'Tot. scont.')) {
    if (!($campi.Name -contains $currentItemName)) {
        [System.Windows.MessageBox]::Show("Header not compliant",'ERROR','Ok','Error') | Out-Null
        Exit
    }
}

$local_array = @() # inizializzo un dataframe in cui raccogliere i dati processati in input

# parsing dei dati in input
$i = 0
$parsebar = ProgressBar
Write-Host -NoNewline "Parsing input file..."
foreach ($record in $rawdata) {
    $i++
    $matches = @()
    $record.'Articolo' -match "(DA|dal)\s+(\d\d/\d\d/\d\d\d\d)\s+(A|al)\s+(\d\d/\d\d/\d\d\d\d)" > $null
    if ($matches[1] -and $matches[3]) { # parso i contratti in abbonamento
        $start_date = [datetime]::parseexact($matches[2], 'dd/MM/yyyy', $null)
        $stop_date = [datetime]::parseexact($matches[4], 'dd/MM/yyyy', $null)

        # calcolo la rata su base mensile (NO canone fisso)
        $conv = ($stop_date - $start_date)
        $mensilita = [math]::Round($conv.TotalDays / 30.42)
        $rata_mensile = [math]::Round(($record.'Tot. scont.' / $mensilita),2)
    } else { # parso i pagamenti una tantum
        $rata_mensile = [math]::Round($record.'Tot. scont.',2)
        if ($record.'Articolo' -match "Storno[\.\/\w\s\n]+del\s+(\d\d/\d\d/\d\d\d\d)") {
            $start_date = $record.'Del'
            $stop_date = 'null'
        } elseif ($record.'Articolo' -match "\s+(\d\d/\d\d/\d\d\d\d)\s+") { # eccezione per pagamenti di cui non e' possibile stabilire se siano una tantum o in abbonamento
            $start_date = 'WARNING'
            $stop_date = 'WARNING'
            [System.Windows.MessageBox]::Show(("Datetime(s) need to be correctly parsed from the following string`n`n{0}" -f ($record.'Articolo')),'WARNING','Ok','Warning') | Out-Null
        } else {
            $start_date = $record.'Del'
            $stop_date = 'null'
        }
    }

    # estraggo il record
    $currentyear = [int32](Get-Date -Format "yyyy") - 1
    $local_hash = [ordered]@{ 
        Tipo        = $record.'Tipo';
        NumDoc      = $record.'Numero Doc';
        Date        = $record.'Del'
        Codice      = $record.'Cod';
        Cliente     = $record.'Cliente';
        CodArt      = $record.'Cod Art';
        Articolo    = $record.'Articolo' -replace "`n"," " -replace "`r",""; # rimuovo i ritorni a capo
        Conto       = $record.'Conto';
        Quantita    = $record.'Qta';
        Importo     = $record.'Tot. scont.';
        Rata        = $rata_mensile;
        Start       = $start_date;
        Stop        = $stop_date;
        ("Apr.{0}" -f ($currentyear)) = 'null';
        ("Mag.{0}" -f ($currentyear)) = 'null';
        ("Giu.{0}" -f ($currentyear)) = 'null';
        ("Lug.{0}" -f ($currentyear)) = 'null';
        ("Ago.{0}" -f ($currentyear)) = 'null';
        ("Set.{0}" -f ($currentyear)) = 'null';
        ("Ott.{0}" -f ($currentyear)) = 'null';
        ("Nov.{0}" -f ($currentyear)) = 'null';
        ("Dic.{0}" -f ($currentyear)) = 'null';
        ("Gen.{0}" -f ($currentyear + 1)) = 'null';
        ("Feb.{0}" -f ($currentyear + 1)) = 'null';
        ("Mar.{0}" -f ($currentyear + 1)) = 'null'
    }

    # aggiungo il record al dataframe
    $local_array += $local_hash

    # avanzamento barra
    $percent = ($i / $rawdata.Count)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $parsebar[2].Text = ("Record {0} out of {1} parsed [{2}%]" -f ($i, $rawdata.Count, $formattato))
    if ($progress -ge 100) {
        $parsebar[1].Value = 100
    } else {
        $parsebar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()  
}
Write-Host -ForegroundColor Green " DONE"
$parsebar[0].Close()

# hashtable per definire il popolamento
$currentyear = [int32](Get-Date -Format "yyyy") -1
$periodi = [ordered]@{ 
    ("Apr.{0}" -f ($currentyear)) = Get-Date -Day 01 -Month 04 -Year $currentyear -Hour 0 -Minute 0;
    ("Mag.{0}" -f ($currentyear)) = Get-Date -Day 01 -Month 05 -Year $currentyear -Hour 0 -Minute 0;
    ("Giu.{0}" -f ($currentyear)) = Get-Date -Day 01 -Month 06 -Year $currentyear -Hour 0 -Minute 0;
    ("Lug.{0}" -f ($currentyear)) = Get-Date -Day 01 -Month 07 -Year $currentyear -Hour 0 -Minute 0;
    ("Ago.{0}" -f ($currentyear)) = Get-Date -Day 01 -Month 08 -Year $currentyear -Hour 0 -Minute 0;
    ("Set.{0}" -f ($currentyear)) = Get-Date -Day 01 -Month 09 -Year $currentyear -Hour 0 -Minute 0;
    ("Ott.{0}" -f ($currentyear)) = Get-Date -Day 01 -Month 10 -Year $currentyear -Hour 0 -Minute 0;
    ("Nov.{0}" -f ($currentyear)) = Get-Date -Day 01 -Month 11 -Year $currentyear -Hour 0 -Minute 0;
    ("Dic.{0}" -f ($currentyear)) = Get-Date -Day 01 -Month 12 -Year $currentyear -Hour 0 -Minute 0;
    ("Gen.{0}" -f ($currentyear + 1)) = Get-Date -Day 01 -Month 01 -Year ($currentyear + 1) -Hour 0 -Minute 0;
    ("Feb.{0}" -f ($currentyear + 1)) = Get-Date -Day 01 -Month 02 -Year ($currentyear + 1) -Hour 0 -Minute 0;
    ("Mar.{0}" -f ($currentyear + 1)) = Get-Date -Day 01 -Month 03 -Year ($currentyear + 1) -Hour 0 -Minute 0
}

# popolo le singole mensilita'
Write-Host -NoNewline "Populating payments..."
$popbar = ProgressBar
for ($i = 0; $i -lt $local_array.Count; $i++) {
    if (($local_array[$i].'Start' -eq 'WARNING') -or ($local_array[$i].'Stop' -eq 'WARNING')) {
        # do_nothing: bypasso l'eccezione gestita sopra
    } elseif (($local_array[$i].'Start' -ne 'null') -and ($local_array[$i].'Stop' -ne 'null')) { # popolamento x abbonamenti
        foreach ($akey in $periodi.Keys) {
            if (($periodi."$akey" -ge $local_array[$i].'Start') -and ($periodi."$akey" -le $local_array[$i].'Stop')) {
                $local_array[$i]."$akey" = $local_array[$i].'Rata'
            }
        }
    } elseif ($local_array[$i].Start -ne 'null') { # popolamento x una tantum
        foreach ($akey in $periodi.Keys) {
            if (($periodi."$akey".Month -eq $local_array[$i].'Start'.Month) -and ($periodi."$akey".Year -eq $local_array[$i].'Start'.Year)) {
                $local_array[$i]."$akey" = $local_array[$i].'Rata'
            }
        }
    }

    # avanzamento barra
    $percent = ($i / $local_array.Count)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $popbar[2].Text = ("Populating {0} out of {1} records [{2}%]" -f ($i, $rawdata.Count, $formattato))
    if ($progress -ge 100) {
        $popbar[1].Value = 100
    } else {
        $popbar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents()
}

Write-Host -ForegroundColor Green " DONE"
$popbar[0].Close()

# produco un csv di output
Write-Host -NoNewline "Writing output file..."
$outfile = "C:\Users\$env:USERNAME\Desktop\ESTRAZIONE.csv"
$header = @($local_array[0].Keys)
$new_string = [system.String]::Join(";", $header)
$new_string | Out-File $outfile -Encoding UTF8
$i = 0
$outbar = ProgressBar
foreach ($item in $local_array) {
    $i++
    $record = @($item.Values)
    $new_string = [system.String]::Join(";", $record)
    $new_string | Out-File $outfile -Encoding UTF8 -Append

    # avanzamento barra
    $percent = ($i / $local_array.Count)*100
    if ($percent -gt 100) {
        $percent = 100
    }
    $formattato = '{0:0.0}' -f $percent
    [int32]$progress = $percent   
    $outbar[2].Text = ("Writing {0} out of {1} lines [{2}%]" -f ($i, $rawdata.Count, $formattato))
    if ($progress -ge 100) {
        $outbar[1].Value = 100
    } else {
        $outbar[1].Value = $progress
    }
    [System.Windows.Forms.Application]::DoEvents() 
}
Write-Host -ForegroundColor Green " DONE"
$outbar[0].Close()

[System.Windows.MessageBox]::Show(("Template created on {0}" -f ($outfile)),'END','Ok','Info') | Out-Null
