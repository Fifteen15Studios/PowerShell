###############################################################################
#
# This script allows users without local admin rights to access the 
#   "ArcGIS Administrator" tool within ArcGIS. The final few lines of the 
#   script may need to be modified based on the installed version of ArcGIS.
#
###############################################################################

$Access = "\\$env:COMPUTERNAME\Users" # user or group to assign access to
$Permission = "Fullcontrol"
# To find all possible values for $Permission, run the following command:
# [system.enum]::getnames([System.Security.AccessControl.RegistryRights])
$Allinherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
$Allpropagation = [system.security.accesscontrol.PropagationFlags]"None"

$AccessRule = New-Object system.security.AccessControl.RegistryAccessRule($Access, $Permission, $AllInherit, $Allpropagation, "Allow")

# Check if User/Group Already has some permissions
if ($GetACL.Access | Where { $_.IdentityReference -eq $Access}) {
	Write-Host "Modifying Permissions For: $Access" -ForeGroundColor Yellow
	$AccessModification = New-Object system.security.AccessControl.AccessControlModification
	$AccessModification.value__ = 2
	$Modification = $False
    # Change permissions accordingly
	$GetACL.ModifyAccessRule($AccessModification, $AccessRule, [ref]$Modification) | Out-Null
} else {
	Write-Host "Adding Permission: $Permission For: $Access"
    # Add permission
	$GetACL.AddAccessRule($AccessRule)
}

# Find path of registry key based on if OS is 32-bit or 64-bits
# Version number - ex: 10.6 - Will be different depending on the version of ArcGIS
#   Adjust this value to match the installed ArcGIS version
if ((Get-CimInstance Win32_OperatingSystem).OSArchitecture -eq "64-bit") {
    $Path = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\ESRI\License10.6"
} else {
    $Path = "HKEY_LOCAL_MACHINE\SOFTWARE\ESRI\License10.6"
}

# Set permissions
Set-Acl -aclobject $GetACL -Path $Path