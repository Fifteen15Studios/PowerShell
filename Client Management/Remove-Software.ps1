###############################################################################
#
# This script does a WMI query for software with a given name and attempts to 
#   remove the software with MSIExec. If there are multiple software packages
#   discovered in the search, it will prompt the user for each piece of 
#   software, asking if the user would like to remove the package.
#
# Input - SoftwareToFind (Required) - A string to search for. This input will
#   be used to search for software which has a name matching this value. A
#   wildcard is added to each side of the input value when the search is 
#   performed.
#
###############################################################################

param(
    [Parameter(Mandatory=$True)]
    $SoftwareToFind
)

cls
"Finding matching software..."
$Software = Get-WmiObject win32_product | Where-Object {$_.name -like "*$SoftwareToFind*"} | sort name
cls

function uninstall($App)
{
    $ActualName = $App.Name
    $ID = $App.IdentifyingNumber
    $Version = $App.Version
    $Answer = Read-Host "Would you like to uninstall" $ActualName $Version"? (Y/[N])"

    if($Answer -ne "" -and $Answer.ToLower()[0] -eq "y")
    {
        "Uninstalling $ActualName silently..."
        Start-Process msiexec -ArgumentList "/qn /x $ID" -Wait
    }
}

# If there is no macth found
if($Software -eq $null)
{
    "$SoftwareToFind is not installed"
}
else
{
    # If there is more than 1 matching application
    if($Software.count -gt 1)
    {
        "Found $($Software.count) matching applications:"
        $Software | ft
    }

    foreach($App in $Software)
    {
        uninstall($App)
    }
}
