param (
    [string]$UserString,                # input string
    [int]$MinimumLength = 10,           # minimum char length
    [switch]$ShuffleBlock = $true,      # split the imput string into blocks of 3 chars and shuffle them
    [switch]$UpperCase = $true,         # capitalize letters
    [switch]$TransLite = $true,         # transliterate
    [switch]$InsDels = $true,           # add single character insertions or deletions
    [switch]$Reverso = $true            # reverts block of 3 chars
)

<#
Name......: PedoMellon.ps1
Version...: 24.10.1
Author....: Dario CORRADA

This script is a password generator: the strings generated should be easy to keep 
in mind but complex enough to decipher, as well

Thx to Marco Motta for his sharing of precious suggestions
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

function Slider {
    param ($form, $x, $y, $min, $max, $defval)

    $obj = New-Object System.Windows.Forms.TrackBar
    $obj.Location = "$x,$y"
    $obj.Width = 200
    $obj.Orientation = "Horizontal"
    $obj.TickStyle = "TopLeft"
    $obj.SetRange($min,$max)
    $obj.LargeChange = $step
    $obj.Value = $defval
    $form.Controls.Add($obj)

    return $obj
}

function TerraForm {
    param (
        [string]$instring, 
        [int]$mlength,
        # methods ("True" if enabled)     
        [string]$shuffle,
        [string]$capitals,
        [string]$tranx,
        [string]$indel,
        [string]$revo,
        # weights: set the frequency by which each method occurs (1/value)
        [int]$wcap = 3,
        [int]$wtra = 3,
        [int]$wind = 20,
        [int]$wrev = 3
    )

    # charsets
    for ($i = 48; $i -lt 58; $i++) {
        $numbers += [char]$i
    }
    for ($i = 65; $i -lt 91; $i++) {
        $maiusc += [char]$i
    }
    for ($i = 97; $i -lt 123; $i++) {
        $letters += [char]$i
    }
    $specials = '!$?*_+#@^=%'

    # lowerize and remove wildcard chars
    $TheSpice = $instring.ToLower()
    $TheSpice = $TheSpice -replace "[\s\\:]+",''

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
            $DnD = Get-Random -Maximum $wcap
            if (($splitted[$i] -match "[a-z]") -and ($DnD -eq 1)) {
                $splitted[$i] = "$($splitted[$i])".ToUpper()
            }
        }
        $TheSpice = -join $splitted
    }
    
    if ($tranx -eq "True") {
        # init transliterate dictionary
        $keys = @('a','b','c','e','g','i','l','o','s','z')
        $vals = @('*','6','(','3','9','!','1','@','$','2')            
        $dict = @{}
        for ($i = 0; $i -lt $keys.Count; $i++) {
            $dict["$($keys[$i])"] = "$($vals[$i])"
        }

        $splitted = $TheSpice.ToCharArray()
        for ($i = 0; $i -lt $splitted.Count; $i++) {
            $akey = "$($splitted[$i])".ToLower()
            $DnD = Get-Random -Maximum $wtra
            if (($dict.ContainsKey($akey)) -and ($DnD -eq 1)) {
                $splitted[$i] = "$($dict[$akey])"
            }
        }
        $TheSpice = -join $splitted
    }

    if ($indel -eq "True") {
        $splitted = $TheSpice.ToCharArray()
        $mutated = @()
        for ($i = 0; $i -lt $splitted.Count; $i++) {
            $DnD = Get-Random -Maximum $wind
            if ($DnD -eq 1) {
                # deletion, do nothing
            } elseif ($DnD -eq 2) {
                # insertion, expand element
                $charset = ($letters + $numbers).ToCharArray()
                $newstring = "$($splitted[$i])"
                $idx = Get-Random -Maximum $charset.Count
                $newstring += "$($charset[$idx])"
                $mutated += $newstring
            } else {
                # no mutation, keep wildtype
                $mutated += "$($splitted[$i])"
            }
        }
        $TheSpice = -join $mutated
    }

    if ($revo -eq "True") {
        $reverted = @()
        for ($i = 0; $i -lt $TheSpice.Length; $i = $i+3) {
            if (($i+2) -lt $TheSpice.Length) {
                $lechunck = $TheSpice.Substring($i,3)
            } else {
                $lechunck = $TheSpice.Substring($i)
            }
            $DnD = Get-Random -Maximum $wrev
            if ($DnD -eq 1) {
                $reverted += $lechunck[-1..-$lechunck.Length] -join ''
            } else {
                $reverted += $lechunck
            }
        }
        $TheSpice = -join $reverted
    }
    
    if ($TheSpice.Length -gt $MinimumLength) {
        $TheSpice = $TheSpice.Substring(0,$MinimumLength)
    } elseif ($TheSpice.Length -lt $MinimumLength) {
        $renmants = $MinimumLength - $TheSpice.Length
        $charset = ($specials + $numbers).ToCharArray()
        for ($i = 0; $i -lt $renmants; $i++) {
            $idx = Get-Random -Maximum $charset.Count
            $TheSpice += "$($charset[$idx])"
        }
    }

    return $TheSpice
}




<# *******************************************************************************
                                    DIALOG
******************************************************************************* #>
if ([string]::IsNullOrEmpty($UserString)) {

    $UserString = "Insert a username..."
    $ThePswd = "Here the password!"

    $continueBrowsing = $true
    while ($continueBrowsing) {
        $TheDialog = FormBase -w 640 -h 360 -text "Pedo Mellon a Minno"

        # disclaimer
        $Disclaimer = Label -form $TheDialog -x 20 -y 10 -w 150 -h 45 -text 'PLEASE NOTE'
        $Disclaimer.Font = [System.Drawing.Font]::new("Arial", 12, [System.Drawing.FontStyle]::Bold)
        $Disclaimer.TextAlign = 'MiddleCenter'
        $Disclaimer.BackColor = 'Red'
        $Disclaimer.ForeColor = 'Yellow'
        $ExLinkLabel = New-Object System.Windows.Forms.LinkLabel
        $ExLinkLabel.Location = New-Object System.Drawing.Size(180,10)
        $ExLinkLabel.Size = New-Object System.Drawing.Size(450,65)
        $ExLinkLabel.Font = [System.Drawing.Font]::new("Arial", 10)
        $ExLinkLabel.Text = @"
The methods adopted herein doesn't ensure cryptographically secure 
randomness. Use such generated string only for temporary purposes. 
Otherwise, feel free to check out the source code and update it as well.
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
        $OptsLabel = Label -form $TheDialog -x 250 -y 120 -h 20 -text "OPTIONS" 
        $OptsLabel.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)
        Label -form $TheDialog -x 20 -y 200 -w 110 -h 20 -text "Characters Length" | Out-Null
        $CharLengthOpt = TxtBox -form $TheDialog -x 130 -y 195 -w 40 -text $MinimumLength
        $ShuffleOpt = CheckBox -form $TheDialog -x 255 -y 140 -checked $ShuffleBlock -text "Block Shuffle"
        $UpperOpt = CheckBox -form $TheDialog -x 255 -y 175 -w 120 -checked $UpperCase -text "Capitalize Letters"
        $TransOpt = CheckBox -form $TheDialog -x 255 -y 210 -w 120-checked $TransLite -text "Transliterate"
        $IndelOpt = CheckBox -form $TheDialog -x 255 -y 245 -checked $InsDels -text "Add Indels"
        $RevertOpt = CheckBox -form $TheDialog -x 255 -y 280 -checked $Reverso -text "Block Revert"

        # trackbars
        $UpperTrack = Slider -form $TheDialog -x 380 -y 175 -min 2 -max 6 -defval 3
        $UpperTrackLabel = Label -form $TheDialog -x 580 -y 185 -text $UpperTrack.Value
        $UpperTrack.add_ValueChanged({
            $SliderValue = $UpperTrack.Value
            $TextString =  $SliderValue
            $UpperTrackLabel.Text = $TextString
        })
        $TransTrack = Slider -form $TheDialog -x 380 -y 220 -min 2 -max 10 -defval 3
        $TransTrackLabel = Label -form $TheDialog -x 580 -y 230 -text $TransTrack.Value
        $TransTrack.add_ValueChanged({
            $SliderValue = $TransTrack.Value
            $TextString =  $SliderValue
            $TransTrackLabel.Text = $TextString
        })
        

        # buttons
        RETRYButton -form $TheDialog -x 20 -y 230 -text "Generate" | Out-Null
        $CopyBut = RETRYButton -form $TheDialog -x 120 -y 230 -text "Copy"
        $CopyBut.DialogResult = [System.Windows.Forms.DialogResult]::IGNORE
        OKButton -form $TheDialog -x 20 -y 270 -w 200 -text "Quit" | Out-Null

        $ButtHead = $TheDialog.ShowDialog()

        if ($ButtHead -eq 'IGNORE') { 
            [System.Windows.Forms.Clipboard]::SetText($PwdBox.Text)
        } elseif ($ButtHead -eq 'RETRY') {
            # cached params
            $UserString     = $UsrBox.Text
            $MinimumLength  = $CharLengthOpt.Text
            $ShuffleBlock   = $ShuffleOpt.Checked
            $UpperCase      = $UpperOpt.Checked
            $TransLite      = $TransOpt.Checked
            $InsDels        = $IndelOpt.Checked
            $Reverso        = $RevertOpt.Checked

            # generate string
            $ThePswd = TerraForm `
                -instring   $UserString `
                -mlength    $MinimumLength `
                -shuffle    $ShuffleBlock.ToString() `
                -capitals   $UpperCase.ToString() `
                -tranx      $TransLite.ToString() `
                -indel      $InsDels.ToString() `
                -revo       $Reverso.ToString()
        } elseif ($ButtHead -eq 'OK') {
            $continueBrowsing = $false
        }
    }

<# *******************************************************************************
                                 HEADLESS
******************************************************************************* #>
} else {
    $ThePswd = TerraForm `
                -instring   $UserString `
                -mlength    $MinimumLength `
                -shuffle    $ShuffleBlock.ToString() `
                -capitals   $UpperCase.ToString() `
                -tranx      $TransLite.ToString() `
                -indel      $InsDels.ToString() `
                -revo       $Reverso.ToString()
    Write-Host "$ThePswd"
}
