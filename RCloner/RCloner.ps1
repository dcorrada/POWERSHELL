<#
Name......: RCloner.ps1
Version...: 24.03.1
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

# get working directory
$fullname = $MyInvocation.MyCommand.Path
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\RCloner\\RCloner\.ps1$" > $null
$repopath = $matches[1]
<# for testing purposes
$repopath = Get-Location
$repopath = $repopath.Path
#>

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

# get remotes list
Write-Host -NoNewline "Looking at remotes list... "
$pinfo = New-Object System.Diagnostics.ProcessStartInfo
$pinfo.FileName = $rclone
$pinfo.RedirectStandardError = $true
$pinfo.RedirectStandardOutput = $true
$pinfo.UseShellExecute = $false
$pinfo.Arguments = "listremotes"
$p = New-Object System.Diagnostics.Process
$p.StartInfo = $pinfo
$p.Start() | Out-Null
$p.WaitForExit()
$stdout = $p.StandardOutput.ReadToEnd()
$stderr = $p.StandardError.ReadToEnd()
Write-Host -ForegroundColor Green "DONE"
$splitted = $stdout.Split("`n")
$adialog = FormBase -w 275 -h ((($splitted.Count-1) * 30) + 120) -text "SELECT A REMOTE"
$they = 20
$choices = @()
foreach ($remote in $splitted) {
    if ($remote -match ":$") {
        if ($they -eq 20) {
            $isfirst = $true
        } else {
            $isfirst = $false
        }
        $choices += RadioButton -form $adialog -x 20 -y $they -checked $isfirst -text $remote
        $they += 30 
    }
}
OKButton -form $adialog -x 75 -y ($they + 10) -text "Ok" | Out-Null
$result = $adialog.ShowDialog()
foreach ($item in $choices) {
    if ($item.Checked) {
        $selected_remote = $item.Text
    }
}

# source and target def
$form = FormBase -w 400 -h 200 -text 'PATHS'
Label -form $form -x 10 -y 20 -w 80 -h 30 -text 'Local source:' | Out-Null
$srcpath = TxtBox -form $form -x 90 -y 20 -w 250 -h 30 -text 'C:\'
Label -form $form -x 10 -y 60 -w 80 -h 30 -text 'Remote target:' | Out-Null
$rmtpath = TxtBox -form $form -x 90 -y 60 -w 250 -h 30 -text '[write here your path]'
OKButton -form $form -x 100 -y 100 -text 'Ok' | Out-Null
$result = $form.ShowDialog()
$source = $srcpath.Text
$target = -join($selected_remote, $rmtpath.Text)

# considering strings of path including spaces
if ($source -match ' ') {
    $source = '"' + $source + '"'
}
if ($target -match ' ') {
    $target = '"' + $target + '"'
}


if ($target -match 'write here your path') {
    [System.Windows.MessageBox]::Show("No defined path for [$selected_remote]",'ERROR','Ok','Error') > $null
    Exit
}
<#
The script process only one volume for each run, defined by $source variable.
Prior to define the destination path ($target) you need to configure the remote
with the interactive 'rclone config' command. The main supported providers are
listed at https://rclone.org/
#>


# logs
$formlist = FormBase -w 400 -h 200 -text 'LOGS'
Label -form $formlist -x 10 -y 20 -w 80 -h 30 -text 'Log level:' | Out-Null
$verbosity = DropDown -form $formlist -x 90 -y 20 -w 250 -h 30 -opts ('DEBUG', 'INFO', 'NOTICE', 'ERROR')
$verbosity.Text = 'NOTICE'
Label -form $formlist -x 10 -y 60 -w 80 -h 30 -text 'Log file:' | Out-Null
$logdia = TxtBox -form $formlist -x 90 -y 60 -w 250 -h 30 -text 'C:\RClone\Melampo.log'
$OKButton = OKButton -form $formlist -x 100 -y 100 -text "Ok"
$result = $formlist.ShowDialog()
$loglevel= $verbosity.Text
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
$dryrun = CheckBox -form $form_panel -checked $false -x 20 -y 20 -w 350 -text "Dry run (do nothing effectively, suggested for testing puposes)"
$links = CheckBox -form $form_panel -checked $true -x 20 -y 50 -w 350 -text "Backup symlink as paths (store links as plain text)"
$hashcheck = CheckBox -form $form_panel -checked $false -x 20 -y 80 -w 350 -text "Disable hash check over SFTP (suggested for very big files)"
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