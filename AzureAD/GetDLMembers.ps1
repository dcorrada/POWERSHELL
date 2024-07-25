<#
Name......: GetDLMembers.ps1
Version...: 22.06.1
Author....: Dario CORRADA

This script will connect to Azure AD and retrieve the members of a Distribution List
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# import the AzureAD module
$ErrorActionPreference= 'Stop'
try {
    Import-Module MSOnline
} catch {
    Install-Module MSOnline -Confirm:$False -Force
    Import-Module MSOnline
}
$ErrorActionPreference= 'Inquire'

# connect to Tenant
Connect-MsolService

# rrawdataroups list and show it
$rawdata = Get-MsolGroup
$local_hash = @{}  
foreach ($item in $rawdata) {
    if ($item.GroupType -eq 'DistributionList') {
        $akey = $item.DisplayName + ' (' + $item.Description + ')'
        $local_hash[$akey] = $item.ObjectId
    }
}

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
foreach ($item in $local_hash.GetEnumerator() | Sort-Object Name) { 
    $DropDown.Items.Add($item.Name)  > $null 
}
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

$groupid = $local_hash[$DropDown.Text]

# retrieve group members
$memberlist = Get-MsolGroupMember -GroupObjectId $groupid

$outfile = "$env:USERPROFILE\Downloads\DLmembers.csv"

$header = 'TYPE;MAIL;FULLNAME'
$header | Out-File $outfile -Encoding UTF8 -Append

foreach ($item in $memberlist) {
    $new_string = "{0};{1};{2}" -f $item.GroupMemberType, $item.EmailAddress, $item.DisplayName
    $new_string | Out-File $outfile -Encoding UTF8 -Append
}

Write-Host "Member List written on [$outfile]"
Pause

