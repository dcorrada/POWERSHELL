<#
Name......: List_Group_Members.ps1
Version...: 19.10.1
Author....: Dario CORRADA

This script list all members of an Active Directory Group
#>

# header 
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# check Active Directory module
if ((Get-Module -Name ActiveDirectory -ListAvailable) -eq $null) {
    $ErrorActionPreference= 'Stop'
    try {
        Get-WindowsCapability -Name RSAT* -Online | Add-WindowsCapability -Online
    }
    catch {
        Write-Host -ForegroundColor Red "Unable to install RSAT"
        Pause
        Exit
    }
    $ErrorActionPreference= 'Inquire'
}

# retrieve AD groups list
$ADgroups = Get-ADGroup -Filter *
$sorted = $ADgroups.Name | Sort-Object

$form = New-Object System.Windows.Forms.Form
$form.Text = "GROUPS LIST"
$form.Size = New-Object System.Drawing.Size(400,230)
$form.StartPosition = 'CenterScreen'

$font = New-Object System.Drawing.Font("Arial", 12)
$form.Font = $font
    
$DropDown = new-object System.Windows.Forms.ComboBox
$DropDown.Location = new-object System.Drawing.Size(10,60)
$DropDown.Size = new-object System.Drawing.Size(350,30)
$DropDown.AutoSize = $true
foreach ($profilo in $sorted) { $DropDown.Items.Add($profilo)  > $null }
$Form.Controls.Add($DropDown)
    
$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
$DropDownLabel.size = new-object System.Drawing.Size(280,30) 
$DropDownLabel.Text = "Select group"
$Form.Controls.Add($DropDownLabel)
    
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Location = New-Object System.Drawing.Point(100,110)
$OKButton.Size = New-Object System.Drawing.Size(75,30)
$OKButton.Text = 'OK'
$OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$form.AcceptButton = $OKButton
$form.Controls.Add($OKButton)
    
$form.Topmost = $true
   
$form.Add_Shown({$DropDown.Select()})
$result = $form.ShowDialog()

$groupname = $DropDown.Text

# retrieve members list of selected group
$ADmembers = Get-ADGroupMember -id $groupname -Recursive 


$outfile = "$env:USERPROFILE\Downloads\ADGroup_members.csv"
"Name;Type;OrganizationalUnit" | Out-File $outfile -Encoding ASCII -Append
foreach ($member in $ADmembers) {
    $member.distinguishedName -match ",OU=([a-zA-Z_\-\.\s0-9]+)," > $null
    $ou = $matches[1]
    
    $new_record = @(
        $member.Name,
        $member.objectClass,
        $ou
    )
    $new_string = [system.String]::Join(";", $new_record)
    $new_string | Out-File $outfile -Encoding ASCII -Append
}