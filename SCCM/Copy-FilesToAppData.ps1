<#
.SYNOPSIS
    This script can be used in SCCM to copy files/folders to the user's AppData folder
 
 .DESCRIPTION
    Instead of using %UserProfile%, this script finds the currently logged on user using WMI
    then uses that username to get to the AppData folder. This is done because the SCCM
    process generally runs as an admin, so %UserProfile% will not be accurate.
    SourcePath is the location of the files being coppied
    AppPath is the path of the destination AFTER \AppData\
 
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
    $SourcePath,
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

$AppData = “C:\Users\$(currentUser)\AppData"

# Source path ending with a \ causes problems sometimes
# If SourcePath ends with \, Remove it
if($SourcePath.EndsWith("\")) {
    $SourcePath=$SourcePath.Substring(0, $SourcePath.Length -1)
}
# If SourcePath is surrounded by quotes, and it ends with \, remove the slash but keep the end quote
elseif($SourcePath.EndsWith("\`"")) {
    $SourcePath="$($SourcePath.Substring(0, $SourcePath.Length - 2))`""
}

# Remove leading \ from AppPath if there is one
if($AppPath.startswith("\")) {
    $AppPath = $AppPath.substring(1)
}

# Compensate for removing \ at the end of source by adding /I
XCOPY $SourcePath "$AppData\$AppPath" /E /Y /H /Q /R /I
