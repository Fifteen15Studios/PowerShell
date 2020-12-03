###############################################################################
#
# This script allows users without local admin rights to access the 
#   "ArcGIS Administrator" tool within ArcGIS.
#
###############################################################################

#requires -runasadmin

$User = "Users" # user or group to assign access to
$Permission = "FullControl"
# To find all possible values for $Permission, run the following command:
# [system.enum]::getnames([System.Security.AccessControl.RegistryRights])
$Allinherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
$Allpropagation = [system.security.accesscontrol.PropagationFlags]"None"

# Create a rule that gives the "Users" group full control over the registry key
$AccessRule = New-Object system.security.AccessControl.RegistryAccessRule($User, $Permission, $AllInherit, $Allpropagation, "Allow")

# If HKCR is not mapped
if(!(Test-path HKCR:)) {
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT | Out-Null
}

# Find path of registry keys based on if OS is 32-bit or 64-bit
if ((Get-CimInstance Win32_OperatingSystem).OSArchitecture -eq "64-bit") {
    $Path = "HKLM:\SOFTWARE\Wow6432Node\ESRI"
    $Path2 = "HKCR:\Wow6432Node\CLSID\{E6BDAA76-4D35-11D0-98BE-00805F7CED21}"
} else {
    $Path = "HKLM:\SOFTWARE\ESRI"
    $Path2 = "HKCR:\CLSID\{E6BDAA76-4D35-11D0-98BE-00805F7CED21}"
}

function Give-Permission($key) {
    # Find the current permissions
    $GetACL = (Get-Item $key).GetAccessControl('Access')

    # Check if User/Group Already has some permissions
    if ($GetACL.Access | Where { $_.IdentityReference -eq $User}) {
	    Write-Host "Modifying Permissions For: $User on $key" -ForeGroundColor Yellow
	    $AccessModification = New-Object system.security.AccessControl.AccessControlModification
	    $AccessModification.value__ = 2
	    $Modification = $False
        # Change permissions accordingly
	    $GetACL.ModifyAccessRule($AccessModification, $AccessRule, [ref]$Modification) | Out-Null
    } else {
	    Write-Host "Adding Permission: $Permission For: $User on $key"
        # Add permission
	    $GetACL.AddAccessRule($AccessRule)
    }

    # Set permissions
    Set-Acl -aclobject $GetACL -Path $key
}

Give-Permission($Path)
Give-Permission($Path2)