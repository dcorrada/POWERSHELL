param (
    [string]$UserString,                # input string
    [int]$MinimumLength = 10,           # minimum char length
    [switch]$ShuffleBlock = $true,      # split the imput string into blocks of 3 chars and shuffle them
    [switch]$UpperCase = $true          # capitalize letters
)

<#
Name......: PedoMellon.ps1
Version...: 24.09.a
Author....: Dario CORRADA

This script is a password generator: the strings generated should be easy to keep 
in mind but complex enough to decipher, as well
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

# just pipe more than single "Split-Path" if the script maps to nested subfolders
$workdir = Split-Path $myinvocation.MyCommand.Definition -Parent | Split-Path -Parent

# graphical stuff
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"


<# *******************************************************************************
                                    METHODS
******************************************************************************* #>

<# 
+++ PLEASE NOTE +++
Get-Random doesn't ensure cryptographically secure randomness. The seed value is 
used for the current command and for all subsequent Get-Random commands in the 
current session until you use SetSeed again or close the session. You can't reset 
the seed to its default value.

Deliberately setting the seed results in non-random, repeatable behavior. It 
should only be used when trying to reproduce behavior, such as when debugging or 
analyzing a script that includes Get-Random commands. Be aware that the seed value 
could be set by other code in the same session, such as an imported module.

PowerShell 7.4 includes Get-SecureRandom, which ensures cryptographically secure 
randomness.

https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-random
#>

function TerraForm {
    param (
        [string]$instring,
        [int]$mlength,     
        [string]$shuffle,
        [string]$capitals
    )

    $TheSpice = $instring.ToLower()

    if ($shuffle -eq "True") {
        $shuffled = @()
        for ($i = 0; $i -lt $TheSpice.Length; $i = $i+3) {
            if (($i+2) -lt $TheSpice.Length) {
                $shuffled += $TheSpice.Substring($i,3)
            } else {
                $shuffled += $TheSpice.Substring($i)
            }
        }
        $shuffled = $shuffled | Sort-Object { Get-Random }
        $TheSpice = -join $shuffled
    }

    if ($capitals -eq "True") {
        $splitted = $TheSpice.ToCharArray()
        for ($i = 0; $i -lt $splitted.Count; $i++) {
            $DnD = Get-Random -Maximum 2
            if (($splitted[$i] -match "[a-z]") -and ($DnD -eq 1)) {
                $splitted[$i] = "$($splitted[$i])".ToUpper()
            }
        }
        $TheSpice = -join $splitted
    }
    
    <#  +++ TODO +++
        * translitterazione
        * inserzioni/delezioni
        * inversione
        * troncare/espandere su $mlength
    #>

    return $TheSpice
}




<# *******************************************************************************
                                    DIALOG
******************************************************************************* #>
if ([string]::IsNullOrEmpty($UserString)) {

    $UserString = "Insert here a username..."
    $ThePswd = "here there is a pswd!"

    $continueBrowsing = $true
    while ($continueBrowsing) {
        $TheDialog = FormBase -w 460 -h 360 -text "PEDO MELLON A MINNO"

        # disclaimer
        $Disclaimer = Label -form $TheDialog -x 20 -y 10 -w 410 -text 'PLEASE NOTE'
        $Disclaimer.Font = [System.Drawing.Font]::new("Arial", 12, [System.Drawing.FontStyle]::Bold)
        $Disclaimer.TextAlign = 'MiddleCenter'
        $Disclaimer.BackColor = 'Red'
        $Disclaimer.ForeColor = 'Yellow'
        $ExLinkLabel = New-Object System.Windows.Forms.LinkLabel
        $ExLinkLabel.Location = New-Object System.Drawing.Size(20,45)
        $ExLinkLabel.Size = New-Object System.Drawing.Size(420,65)
        $ExLinkLabel.LinkColor = "Blue"
        $ExLinkLabel.Font = [System.Drawing.Font]::new("Arial", 10)
        $ExLinkLabel.Text = @"
The methods adopted herein doesn't ensure cryptographically 
secure randomness. Use such generated string only for temporary 
purposes. Otherwise, feel free to check out the source code and 
update it as well.
"@        
        $ExLinkLabel.add_Click({[system.Diagnostics.Process]::start("https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/get-random")})
        $TheDialog.Controls.Add($ExLinkLabel)

        # string area
        $UsrBox = TxtBox -form $TheDialog -x 20 -y 120 -h 35 -w 200 -text $UserString
        $UsrBox.Font = [System.Drawing.Font]::new("Arial", 11)
        $PwdBox = TxtBox -form $TheDialog -x 20 -y 160 -h 35 -w 200 -text $ThePswd
        $PwdBox.ReadOnly = $true
        $PwdBox.Font = [System.Drawing.Font]::new("Courier New", 11)

        # methods area
        $OptsLabel = Label -form $TheDialog -x 250 -y 110 -h 20 -text "OPTIONS" 
        $OptsLabel.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)
        Label -form $TheDialog -x 250 -y 135 -w 120 -h 20 -text "Minimum Char Length" | Out-Null
        $CharLengthOpt = TxtBox -form $TheDialog -x 370 -y 132 -w 30 -text $MinimumLength
        $ShuffleOpt = CheckBox -form $TheDialog -x 250 -y 155 -checked $ShuffleBlock -text "Block Shuffle"
        $UpperOpt = CheckBox -form $TheDialog -x 250 -y 185 -checked $UpperCase -text "Capitalize Letters"
        $TransOpt = CheckBox -form $TheDialog -x 250 -y 215 -enabled $false -text "Transliterate"
        $IndelOpt = CheckBox -form $TheDialog -x 250 -y 245 -enabled $false -text "Add Indels"
        $RecertOpt = CheckBox -form $TheDialog -x 250 -y 275 -enabled $false -text "Block Revert"

        # buttons
        RETRYButton -form $TheDialog -x 20 -y 200 -text "Generate" | Out-Null
        $CopyBut = RETRYButton -form $TheDialog -x 120 -y 200 -text "Copy"
        $CopyBut.DialogResult = [System.Windows.Forms.DialogResult]::IGNORE
        OKButton -form $TheDialog -x 20 -y 270 -w 200 -text "Quit" | Out-Null

        $ButtHead = $TheDialog.ShowDialog()

        if ($ButtHead -eq 'IGNORE') { 
            [System.Windows.Forms.Clipboard]::SetText($PwdBox.Text)
        } elseif ($ButtHead -eq 'RETRY') {
            # cached params
            $UserString     = $UsrBox.Text
            $MinimumLength  = $CharLengthOpt.Text
            $ShuffleBlock   =  $ShuffleOpt.Checked
            $UpperCase      = $UpperOpt.Checked

            # generate string
            $ThePswd = TerraForm `
                -instring   $UserString `
                -mlength    $MinimumLength `
                -shuffle    $ShuffleBlock.ToString() `
                -capitals   $UpperCase.ToString()
        } elseif ($ButtHead -eq 'OK') {
            $continueBrowsing = $false
        }
    }

<# *******************************************************************************
                                 HEADLESS
******************************************************************************* #>
} else {
    <# generatore senza dialog grafica, per integrare lo script in altri #>
}