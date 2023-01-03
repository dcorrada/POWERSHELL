<#
Name......: Bluetooth_fix.ps1
Version...: 23.1.1
Author....: Dario CORRADA

This script performs a bluetooth troubleshooting protocol described in
https://www.thewindowsclub.com/bluetooth-devices-not-showing-windows

See also this post about fixing Bluetooth issues:
https://answers.microsoft.com/en-us/windows/forum/windows_10-networking-winpc/unable-to-remove-bluetooth-device-on-windows-10/ea6da83d-583e-4b80-8714-367510879f07
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
$fullname -match "([a-zA-Z_\-\.\\\s0-9:]+)\\Troubleshooting\\Bluetooth_fix\.ps1$" > $null
$workdir = $matches[1]

# header 
$ErrorActionPreference= 'SilentlyContinue'
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy Bypass -Force
Write-Host "ExecutionPolicy Bypass" -fore Green
$ErrorActionPreference= 'Inquire'
$WarningPreference = 'SilentlyContinue'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
Import-Module -Name "$workdir\Modules\Forms.psm1"

Write-Host -ForegroundColor Blue "`nSTEP 00: Unjoin device(s)"
# get a list of device(s) to be disabled
$btdevices = @{}
foreach ($item in (Get-PnpDevice -Class Bluetooth)) {
    if ($item.HardwareID -imatch 'dev') {
        $ErrorActionPreference= 'Stop'
        try {
            $hex = [uInt64]('0x{0}' -f $item.HardwareID[0].Substring(12))
        }
        catch {
            $hex = 'na'
        }
        $ErrorActionPreference= 'Inquire'
        $btdevices[$item.FriendlyName] = @{
            'Name'     = $item.FriendlyName
            'Status'   = $item.Status
            'Instance' = $item.InstanceID
            'Address'  = $hex
        }
    }
}
$adialog = FormBase -w 350 -h ((($btdevices.Keys.Count-1) * 30) + 200) -text "JOINED DEVICES"
Label -form $adialog -x 10 -y 20 -w 150 -h 30 -text 'Select devices to unjoin:' | Out-Null
$they = 50
$choices = @()
foreach ($item in ($btdevices.Keys | Sort-Object)) {
    $choices += CheckBox -form $adialog -x 20 -y $they -checked $False -text $item
    $they += 30
}
OKButton -form $adialog -x 100 -y ($they + 20) -text "Ok" | Out-Null
$result = $adialog.ShowDialog()
# disable and unpair devices according to the two answer posted on
# https://stackoverflow.com/questions/53642702/how-to-connect-and-disconnect-a-bluetooth-device-with-a-powershell-script-on-win
$Source = @"
   [DllImport("BluetoothAPIs.dll", SetLastError = true, CallingConvention = CallingConvention.StdCall)]
   [return: MarshalAs(UnmanagedType.U4)]
   static extern UInt32 BluetoothRemoveDevice(IntPtr pAddress);
   public static UInt32 Unpair(UInt64 BTAddress) {
      GCHandle pinnedAddr = GCHandle.Alloc(BTAddress, GCHandleType.Pinned);
      IntPtr pAddress     = pinnedAddr.AddrOfPinnedObject();
      UInt32 result       = BluetoothRemoveDevice(pAddress);
      pinnedAddr.Free();
      return result;
   }
"@
$BTR = Add-Type -MemberDefinition $Source -Name "BTRemover"  -Namespace "BStuff" -PassThru
foreach ($item in $choices) {
    if ($item.Checked) {
        $devname = $item.Text
        Write-Host -ForegroundColor Yellow "`n*** $devname ***"
        Write-Host -NoNewline "Disabling... "
        $ErrorActionPreference= 'Stop'
        try {
            Disable-PnpDevice -InstanceId $btdevices[$devname].Instance -Confirm:$false
            Start-Sleep 1
            Write-Host -ForegroundColor Green 'OK'
        }
        catch {
            Write-Host -ForegroundColor Red 'KO'
            Write-Output "Error: $($error[0].ToString())`n"
            Pause
        }
        $ErrorActionPreference= 'Inquire'
        Write-Host -NoNewline "Unpairing... "
        if ($btdevices[$devname].Address -eq 'na') {
            Write-Host -ForegroundColor Yellow 'skipped'
        } else {
            $success = $BTR::Unpair($btdevices[$devname].Address)
            if (!$success) {
                Start-Sleep 1
                Write-Host -ForegroundColor Green 'OK'
            } else {
                Write-Host -ForegroundColor Red 'KO'
            }
        }
    }
}
Pause

# run Hardware Troubleshooter
Write-Host -ForegroundColor Blue "`nSTEP 01: Hardware Troubleshooter"
$answ = [System.Windows.MessageBox]::Show("Run HW Troubleshooter?",'CMD','YesNo','Info')
if ($answ -eq "Yes") {
    $bins = 'C:\Windows\System32\msdt.exe'
    $StagingArgumentList = '-id DeviceDiagnostic'
    
    $ErrorActionPreference= 'Stop'
    try {
        Start-Process -Wait -FilePath $bins -ArgumentList $StagingArgumentList -NoNewWindow
    }
    catch {
        Write-Output "`nError: $($error[0].ToString())"
        Pause
        Exit
    }
    $ErrorActionPreference= 'Inquire'   
}
Pause

# restart Bluetooth SupportService
Write-Host -ForegroundColor Blue "`nSTEP 02: Restart Bluetooth Support"
Write-Host -NoNewline "Restarting service... "
Restart-Service -Name 'bthserv' -Force
if ((Get-Service "bthserv").Status -eq 'Running') {
    Start-Sleep 2
    Write-Host -ForegroundColor Green 'OK'
} else {
    Write-Host -ForegroundColor Red 'KO (not running)'
}
Pause


# run Device Manager
Write-Host -ForegroundColor Blue "`nSTEP 03: Manage Bluetooth drivers"
$answ = [System.Windows.MessageBox]::Show("Run Device Manager?",'CMD','YesNo','Info')
if ($answ -eq "Yes") {
    & devmgmt.msc 
}

