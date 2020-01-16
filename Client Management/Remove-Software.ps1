param(
    [Parameter(Mandatory=$True)]
    $SoftwareToFind
)

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
        msiexec /qn /x $ID
    }

}


# If there is no macth found
if($Software -eq $null)
{
    "$SoftwareToFind is not installed"
}
# If there is more than 1 matching application
elseif($Software.count -gt 1)
{
    "Found $($Software.count) matching applications:"
    $Software | ft

    foreach($App in $Software)
    {
        uninstall($App)
    }
}
# If only 1 application is found
else
{
    uninstall($Software)
}
