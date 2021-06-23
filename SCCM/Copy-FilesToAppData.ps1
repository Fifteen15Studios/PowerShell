<#
.SYNOPSIS
    This script can be used in SCCM to copy files/folders to the user's AppData folder
 
 .DESCRIPTION
    Instead of using %UserProfile%, this script finds the currently logged on user using WMI
    then uses that username to get to the AppData folder. This is done because the SCCM
    process generally runs as an admin, so %UserProfile% will not be accurate.
 
.LINK
    

.NOTES
    Version: 1.0
    DateCreated: 2021-06-23
 
.EXAMPLE
    Copy-FilesToAppData "%~dp0folderToCopy\Subfolder" "Roaming\MyApp\MyFolder"

    Since you are likely using this with SCCM, you will probably want to use %~dp0 to start the FilePath variable
    to point it to the current directory of the running script.
#>

param(
    [parameter(Mandatory=$True)]
    $FilePath,
    [parameter(Mandatory=$True)]
    $AppPath
)

Function CurrentUser {
     $LoggedInUser = get-wmiobject win32_computersystem | select username
     $LoggedInUser = [string]$LoggedInUser
     $LoggedInUser = $LoggedInUser.split(“\”)
     $LoggedInUser = $LoggedInUser.split(“}”)
     $LoggedInUser = $LoggedInUser[1]
     Return $LoggedInUser
}

$user = CurrentUser

$AppData = “C:\Users\” + $user + “\AppData"
XCOPY $FilePath "$AppData\$AppPath" /E /Y /H /Q /R