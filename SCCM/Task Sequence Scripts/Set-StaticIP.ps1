# This script sets a static IP address as part of an OSD task sequence.
#
# Generally, this should probably be the last step of your task sequence
# unless the static is required to connect at the location where the 
# image is being applied, in which case it may be moved to the first step.
#
# The script grabs the IP Address, subnet mask and default gateway from 
# task sequence variables named IPAddress, Subnet, and Gateway respectively.
# These task sequence variables must be set before calling this script.

#Import Task Sequence environment
$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment

#Change these to your DNS servers
$DNS1 = "192.168.0.101"
$DNS2 = "192.168.0.102"

#Pull information from TS Variables
$strIPAddress = $tsenv.Value("IPAddress")
$strSubnet = $tsenv.Value("Subnet")
$strGateway = $tsenv.Value("Gateway")

#Set IP information
$NetworkConfig = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IpEnabled = 'True'"
$NetworkConfig.EnableStatic($strIPAddress, $strSubnet)
$NetworkConfig.SetGateways($strGateway, 1)
$NetworkConfig.SetDNSServerSearchOrder(@($DNS1, $DNS2))
