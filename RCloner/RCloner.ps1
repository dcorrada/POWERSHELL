<#
Name......: RCloner.ps1
Version...: 22.04.1
Author....: Dario CORRADA

This script performs a total synchronized backup from your local device to a 
remote storage, using the RClone client. For usage see https://rclone.org/docs/
#>

# elevated script execution with admin privileges
$currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
$testadmin = $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
if ($testadmin -eq $false) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    exit $LASTEXITCODE
}

# setting script execution policy
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
$ErrorActionPreference= 'Inquire'

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\RCloner\\RCloner\.ps1$" > $null
$repopath = $matches[1]

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$repopath\Modules\Forms.psm1"

# check out
$rclone = 'C:\RClone\rclone.exe' # binary of the RClone client
if (!(Test-Path $rclone)) {
    [System.Windows.MessageBox]::Show("RClone client [$rclone] not found",'ERROR','Ok','Error') > $null
    Exit
}
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select filtering rules file"
$OpenFileDialog.initialDirectory = 'C:\RClone'
$OpenFileDialog.filename = 'filter.list'
$OpenFileDialog.filter = 'Plain text file | *.*'
$OpenFileDialog.ShowDialog() | Out-Null
$filter_list = $OpenFileDialog.filename # plain text file containing filtering rules


# source and target def
$form = FormBase -w 400 -h 200 -text 'PATHS'
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(80,30)
$label.Text = 'Local Source:'
$form.Controls.Add($label)
$srcpath = New-Object System.Windows.Forms.TextBox
$srcpath.Location = New-Object System.Drawing.Point(90,20)
$srcpath.Size = New-Object System.Drawing.Size(250,30)
$srcpath.Text = 'C:\'
$form.Controls.Add($srcpath)
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(10,60)
$label2.Size = New-Object System.Drawing.Size(80,30)
$label2.Text = 'Remote target:'
$form.Controls.Add($label2)
$rmtpath = New-Object System.Windows.Forms.TextBox
$rmtpath.Location = New-Object System.Drawing.Point(90,60)
$rmtpath.Size = New-Object System.Drawing.Size(250,30)
$rmtpath.Text = 'Melampo:/media/CAMUS/BACKUP_AGM/'
$form.Controls.Add($rmtpath)
$OKButton = New-Object System.Windows.Forms.Button
OKButton -form $form -x 100 -y 100 -text "Ok"
$form.Topmost = $true
$result = $form.ShowDialog()
$source = $srcpath.Text
$target = $rmtpath.Text
<#
The script process only one volume for each run, defined by $source variable.
Prior to define the destination path ($target) you need to configure the remote
with the interactive 'rclone config' command. The main supported providers are
listed at https://rclone.org/
#>


# logs
$formlist = FormBase -w 400 -h 200 -text 'LOGS'
$DropDownLabel = new-object System.Windows.Forms.Label
$DropDownLabel.Location = new-object System.Drawing.Size(10,20) 
$DropDownLabel.size = new-object System.Drawing.Size(80,30) 
$DropDownLabel.Text = "Log level:"
$formlist.Controls.Add($DropDownLabel)
$DropDown = new-object System.Windows.Forms.ComboBox
$DropDown.Location = new-object System.Drawing.Size(90,20)
$DropDown.Size = new-object System.Drawing.Size(250,30)
foreach ($elem in ('DEBUG', 'INFO', 'NOTICE', 'ERROR')) {
    $DropDown.Items.Add($elem)  > $null
}
$DropDown.Text = 'NOTICE'
$formlist.Controls.Add($DropDown)
$filelabel = New-Object System.Windows.Forms.Label
$filelabel.Location = New-Object System.Drawing.Point(10,60)
$filelabel.Size = New-Object System.Drawing.Size(80,30)
$filelabel.Text = 'Log file:'
$formlist.Controls.Add($filelabel)
$logdia = New-Object System.Windows.Forms.TextBox
$logdia.Location = New-Object System.Drawing.Point(90,60)
$logdia.Size = New-Object System.Drawing.Size(250,30)
$logdia.Text = 'C:\RClone\Melampo.log'
$formlist.Controls.Add($logdia)
OKButton -form $formlist -x 100 -y 100 -text "Ok"
$formlist.Add_Shown({$DropDown.Select()})
$result = $formlist.ShowDialog()
$loglevel= $DropDown.Text
$logfile = $logdia.Text

<# FLAGS
--dry-run
    do a trial run with no permanent changes. Use this to see what rclone would
    do without actually doing it. 

--filter-from <C:\a_path\filter.list>
    specify  a named file containing a list of remarks and pattern rules. 
    Include rules start with '+' and exclude rules with '-', rules are processed
    in the order they are defined. See also https://rclone.org/filtering/ 

--links
    rclone will copy symbolic links from the local storage, and store them as 
    text files, with a '.rclonelink' suffix in the remote storage.

--log-level <STRING>
    defines the verbosity for logs:

    DEBUG   outputs lots of debug info, useful for bug reports and really 
            finding out what rclone is doing.

    INFO    outputs information about each transfer and prints stats once 
            a minute by default.

    NOTICE  is the default log level, it outputs warnings and significant 
            events.

    ERROR   only outputs error messages.

--sftp-disable-hashcheck
    disable the execution of SSH commands to determine if remote file 
    hashing is available.
#>
$form_panel = FormBase -w 400 -h 200 -text "BEHAVIOUR"
$dryrun = CheckBox -form $form_panel -checked $false -x 20 -y 20 -text "Dry run (do nothing effectively, suggested for testing puposes)"
$links = CheckBox -form $form_panel -checked $true -x 20 -y 50 -text "Backup symlink as paths (store links as plain text)"
$hashcheck = CheckBox -form $form_panel -checked $false -x 20 -y 80 -text "Disable hash check over SFTP (suggested for very big files)"
OKButton -form $form_panel -x 100 -y 120 -text "Ok"
$result = $form_panel.ShowDialog()
$flags = ('--progress', "--log-level $loglevel", "--log-file $logfile", "--filter-from $filter_list")
if ($dryrun.Checked) {
    $flags += '--dry-run'
}
if ($links.Checked) {
    $flags += '--links'
}
if ($hashcheck.Checked) {
    $flags += '--sftp-disable-hashcheck'
}

if (Test-Path $logfile) { 
    # the log file will be overwritten each run
    Remove-Item $logfile -Force
}
$StagingArgumentList = 'sync {0} {1} {2}' -f $source, $target, ($flags -join ' ')



# job launch
$ErrorActionPreference= 'Stop'
try {
    Start-Process -Wait -FilePath $rclone -ArgumentList $StagingArgumentList -NoNewWindow
}
catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    Exit
}
$ErrorActionPreference= 'Inquire'

# summary
$answ = [System.Windows.MessageBox]::Show("Your job has been accomplished.`nDo you want to see the log?","THAT'S ALL",'YesNo','Info')
if ($answ -eq "Yes") {    
    notepad $logfile
}