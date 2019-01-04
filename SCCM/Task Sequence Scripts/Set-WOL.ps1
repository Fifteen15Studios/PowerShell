# Turns on Wake on LAN in HP computers.
#
# I noticed that the options for HP WoL BIOS Settings were not always named the
# same way, but that the option that we want is always the last option. This
# is sometimes called "boot to disk" or "follow boot order" but it was always
# the last in the list. 
#
# This script uses the HP BIOS Config Utility
# (http://ftp.hp.com/pub/caps-softpaq/cmit/HP_BCU.html)
# to set the WoL setting to be the last option in the list.
#
# This must be used as part of a package, and that package should point to
# the files for the BCU in order to work.

$Setting = "Wake On LAN"

#Find possible Wake on LAN settings
$instance = get-ciminstance -classname hp_biossetting -namespace "root\hp\instrumentedbios" | where name -eq $Setting
#Get the last one in the list
$choice = ($instance.possiblevalues)[-1]

#Set the setting to the proper value
if((Get-WmiObject Win32_OperatingSystem).OSArchitecture -eq "64-bit")
{
    .\BiosConfigUtility64.exe /setvalue:"$Setting","$choice"
}
else
{
    .\BiosConfigUtility.exe /setvalue:"$Setting","$choice"
}
