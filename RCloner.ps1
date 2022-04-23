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

# graphical stuff
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework

# check out
$rclone = 'C:\RClone\rclone.exe' # binary of the RClone client
$filter_list = 'C:\RClone\filter.list' # plain text file containing filtering rules
$logfile = 'C:\RClone\Melampo.log' # summary log of the running job
if (!(Test-Path $rclone)) {
    [System.Windows.MessageBox]::Show("RClone client [$rclone] not found",'ERROR','Ok','Error') > $null
    Exit
} elseif (!(Test-Path $filter_list)) {
    [System.Windows.MessageBox]::Show("Filtering rules [$filter_list] not found",'ERROR','Ok','Error') > $null
    Exit
}

# jobrun
$ErrorActionPreference= 'Stop'
try {
    $source = 'C:\'
    $target = 'Melampo:/media/CAMUS/BACKUP_AGM/'
<#
The script process only one volume for each run, defined by $source variable.
Prior to define the destination path ($target) you need to configure the remote
with the interactive 'rclone config' command. The main supported providers are
listed at https://rclone.org/
#>

    $flags = ('--progress', '--links', '--log-level NOTICE', "--log-file $logfile", "--filter-from $filter_list") #, '--dry-run')
<#
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
#>

    if (Test-Path $logfile) { 
        # the log file will be overwritten each run
        Remove-Item $logfile -Force
    }

    $StagingArgumentList = 'sync {0} {1} {2}' -f $source, $target, ($flags -join ' ')
    Start-Process -Wait -FilePath $rclone -ArgumentList $StagingArgumentList -NoNewWindow
}
catch {
    Write-Output "`nError: $($error[0].ToString())"
    Pause
    Exit
}
$ErrorActionPreference= 'Inquire'