Once PATROL is installed launch the script create_DB.ps1. Two files will be created:

	* crypto.key        -> the crypto key
	* PatrolDB.csv.AES  -> the encrypted DB


N.B.: PATROL check only AD credential, it is recommended to join your computer to a local domain.

In order to protect your script using PATROL access grants:

1) Add, on the header of your script, the following code:

	# grant access
	Import-Module -Name "[installation_path]\Modules\Patrol.psm1"
	$credential = Patrol -scriptname [script_name]
   
   The "Patrol" function return $credential object, containg username/password credentials
   Password is encrypted accordoing to SecureString method 

2) Launch UpdateDB.ps1 script and click on "Add record" option to add the following values
   [script_name] and username to be granted

To further protect your scripts it is suggested to compile them using a tool like PS2EXE:

	https://gallery.technet.microsoft.com/scriptcenter/PS2EXE-Convert-PowerShell-9e4e07f1