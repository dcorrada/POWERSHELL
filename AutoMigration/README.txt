These scripts performs user's data migration from a source computer to a destination computer.
In order to perform data migration do the following steps:

1) On the source computer launch LocalAdmin-Cshared.ps1

2) On the destination computer create account(s) identical of the source computer and access once for each of them

3) On the destination computer launch Automigration.ps1

4) Once data migration has terminated, on the source computer disable sharing on the volume C: (Unshare.ps1) 

TODO:
* support for virtual machines
* support for domain users