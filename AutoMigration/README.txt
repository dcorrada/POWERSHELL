These scripts allow complete data migration from a source computer to a destination computer.
In order to perform data migration do the following steps:

1) On the source computer launch LocalAdmin-Cshared.ps1

2) On the destination computer create an account identical to the source computer (or access with the same domain account)

3) On the destination computer install the applications you want to migrate profiles

2) On the destination computer launch Automigration.ps1

3) Once data migration has terminated, on the source computer disable sharing on the volume C: (Unshare.ps1) 