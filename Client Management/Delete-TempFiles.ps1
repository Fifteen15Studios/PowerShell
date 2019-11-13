###############################################################################
#
# Deletes files from various locations used for temporary storage
#
###############################################################################

#Requires -runasadministrator
function Delete-TempFiles()
{
    $OSVersion = [environment]::OSVersion.Version

    $tempFolders = @(
    "C:\Windows\Temp\*", # Default Windows temp location
    "C:\windows\ccmcache\*" # Default SCCM temp location
    )

    # Windows XP user temp
    if($OSVersion.Major -eq 5 -and $OSVersion.Minor -ge 1)
    {
        $tempFolders += "C:\Documents and Settings\*\Local Settings\temp\*"
    }
    # Windows Vista+ user temp
    elseif($OSVersion.Major -ge 6)
    {
        $tempFolders += "C:\Users\*\Appdata\Local\Temp\*"
    }

    Remove-Item $tempFolders -recurse -force -ErrorAction SilentlyContinue
}
