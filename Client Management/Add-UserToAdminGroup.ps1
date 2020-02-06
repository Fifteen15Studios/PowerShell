###############################################################################
#
# This script adds a user to the Administrators group on a PC. This requires 
#    that the script be run as a user that already has administrator privileges
#    on the PC.
#
# Parameters:
#    ComputerName - Name of the computer on which to add the user
#    DomainName - Name of the domain to which the user belongs
#    UserName - uUsername of the user that should be added to the Administrators
#        group.
#
###############################################################################

param(
    [String]$ComputerName,
    [String]$DomainName,
    [parameter(mandatory=$true)]
    [String]$UserName
)

# If computer name not present, use local host name
if($ComputerName -eq $null -or $ComputerName -eq "")
{$ComputerName = $env:COMPUTERNAME}

$AdminGroup = [ADSI]"WinNT://$ComputerName/Administrators,group"

# If domain name not present, use computer name
if($DomainName -eq $null -or $DomainName -eq "")
{$DomainName = $ComputerName}

$User = [ADSI]"WinNT://$DomainName/$UserName,user"
$AdminGroup.Add($User.Path)
