##################################################################################
#
#
#  Script name: Set-Permission.ps1
#  Author:      goude@powershell.nu
#  Homepage:    www.powershell.nu
#
# This is a script that I "stole" from elsewhere and modified to make it my own.
#
# This script will set a user or group's access permissions on a file, folder or
# registry item.
#
# Run 'Set-Permission.ps1 -help' for more information
##################################################################################

param ([string]$Path, [string]$Access, [string]$Permission = ("Modify"), [switch]$help)

function GetHelp() {

$HelpText = @"

DESCRIPTION:
NAME: Set-Permission.ps1
Sets Permissions for User on a Folder or Registry Key.
Creates folder/key if not exist.

PARAMETERS: 
-Path			Folder/Key to Create or Modify (Required)
-Access			User or group who should have access (Required)
-Permission		Specify Permission for User, Default set to Modify (Optional)
-help			Prints the HelpFile (Optional)

SYNTAX:
.\Set-Permission.ps1 -Path <path> -Access [Domain\]<UserName> [-Permission <Access Level>]

EXAMPLE 1:
.\Set-Permission.ps1 -Path C:\Folder\NewFolder -Access PowerUsers -Permission FullControl

Creates the folder C:\Folder\NewFolder if it doesn't exist.
Sets Full Control for local PowerUsers

EXAMPLE 2:
.\Set-Permission.ps1 -Path "C:\Program Files\NewFolder" -Access Domain\UserName

Creates the folder C:\Program Files\NewFolder if it doesn't exist.
Sets Modify (Default Value) for Domain\UserName

EXAMPLE 3:
.\Set-Permission.ps1 -Path "HKLM:\Software\SomeSoftware\RegKey" -Access Users -Permission FullControl

Sets Full Control for Users on HKEY_LOCAL_MACHINE\Software\SomeSoftware\RegKey


Below Are Available Values for -Permission

FILE/FOLDER:

"@
$HelpText

[system.enum]::getnames([System.Security.AccessControl.FileSystemRights])
"
REGISTRY:
"
[system.enum]::getnames([System.Security.AccessControl.RegistryRights])

}

function CreateFolder ([string]$Path) {       

	# Check if the Folder/File Exists
	if (Test-Path $Path) {
		Write-Host "$Path Already Exists" -ForeGroundColor Yellow
	} else {
		Write-Host "Creating folder $Path" -Foregroundcolor Green
		New-Item -Path $Path -type directory | Out-Null
	}
}

function SetAcl ([string]$Path, [string]$Access, [string]$Permission) {

	# Get ACL on Folder
	$GetACL = (Get-Item $Path).GetAccessControl('Access')

	# Set up AccessRule
	$Allinherit = [system.security.accesscontrol.InheritanceFlags]"ContainerInherit, ObjectInherit"
    $NoInherit = [system.security.accesscontrol.InheritanceFlags]::None
	$Allpropagation = [system.security.accesscontrol.PropagationFlags]"None"

    #If it's a registry key, setup registry access rule
    if($path.StartsWith("HKLM:") -or $path.StartsWith("HKCU:") -or $path.StartsWith("HKCR:") -or $path.StartsWith("HKU:"))
    {
        #If ket is in HKCR, map drive to that part of the registry
        if($path.StartsWith("HKCR:"))
        {
            New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT
        }
        #If ket is in HKU, map drive to that part of the registry
        if($path.StartsWith("HKU:"))
        {
            New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS
        }
        $AccessRule = New-Object system.security.AccessControl.RegistryAccessRule($Access, $Permission, $AllInherit, $Allpropagation, "Allow")
    }
    #If it's a file or folder, create file system access rule
    else
    {
        $file = get-item $path

        #If it's a folder, set inheritance
        if($file.PSIsContainer)
        {
	        $AccessRule = New-Object system.security.AccessControl.FileSystemAccessRule($Access, $Permission, $AllInherit, $Allpropagation, "Allow")
        }
        #If file, can't set inheritance
        else
        {
            $AccessRule = New-Object system.security.AccessControl.FileSystemAccessRule($Access, $Permission, $NoInherit, $Allpropagation, "Allow")
        }
    }

	# Check if Access Already Exists
	if ($GetACL.Access | Where { $_.IdentityReference -eq $Access}) {

		Write-Host "Modifying Permissions For: $Access" -ForeGroundColor Yellow

		$AccessModification = New-Object system.security.AccessControl.AccessControlModification
		$AccessModification.value__ = 2
		$Modification = $False
		$GetACL.ModifyAccessRule($AccessModification, $AccessRule, [ref]$Modification) | Out-Null
	} else {

		Write-Host "Adding Permission: $Permission For: $Access"

		$GetACL.AddAccessRule($AccessRule)
	}

	Set-Acl -aclobject $GetACL -Path $Path

	Write-Host "Permission: $Permission Set For: $Access" -ForeGroundColor Green
}

if ($help) { GetHelp }

if ($Path -AND $Access -AND $Permission) { 
	CreateFolder $Path 
	SetAcl $Path $Access $Permission
}
