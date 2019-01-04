# Purpose:
# Searches through your SCCM environment to find driver packs that are not 
# referenced by any Task Sequence. If any such driver packs are found, the pack
# will be deleted out of SCCM, and the source files will also be deleted from
# the location where they are stored.
#
# Parameters:
#
# WhatIf [switch] - Displays the name, location, and size of the ununsed driver
#    packs, and a total size of all of the driver packs, but does not delete
#    anything from SCCM or any files.

[CmdletBinding()]
param(
    [switch]$WhatIf
)

function Get-FolderSize($path)
{
    #"{0:N2} MB" -f 
    ((Get-ChildItem $path -Recurse | Measure-Object -Property Length -Sum).Sum / 1Mb)
}

# Import SCCM Module
Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1)

# Set location to SCCM site
$curLocation = (Get-Location).path
$SiteCode = Get-PSDrive -PSProvider CMSITE
set-location "$($SiteCode):"

# Get all Task Sequences
$TaskSequences = Get-CMTaskSequence
# Get all Driver Packages
$DriverPackages = Get-CMDriverPackage

$Used = @()

# Find all used Driver Packages
foreach($TS in $TaskSequences)
{
    $References = $TS.References
    # Loop through Task Sequence References
    foreach($Reference in $References)
    {
        #If reference is a driver package
        if($DriverPackages.PackageID -contains $Reference.Package)
        {
            $ID = (Get-CMDriverPackage -id $Reference.Package).PackageID
            #If not yet in list, add to list
            if($Used -notcontains $ID)
            {
                $Used += $ID
            }
        }
    }
}

# Turn Driver Packages variable into a list
$Unused = [System.Collections.ArrayList]@()
foreach($DP in $DriverPackages)
{
    $Unused.add($DP) | Out-Null
}

# Remove all used Driver Packages
foreach($DP in $DriverPackages)
{
    if($DP.packageID -in $Used)
    {
        $Unused.Remove($DP)
    }
}
# $Unused now only contains Driver Packages that are not in use

if($WhatIf)
{
    #Cannot access the files from CM-Site. 
    #Set location back to an area that can access these files
    Set-Location C:
    #Find size of each package
    foreach($DP in $Unused)
    {
        $Size = Get-FolderSize($DP.PkgSourcePath)
        $DP | Add-Member -MemberType NoteProperty -Name "Size (MB)" -Value $Size -Force
    }

    "Unused Driver Packages:"
    $Unused | Ft -Property Name,PkgSourcePath,"Size (MB)"
    $size=0
    foreach($package in $Unused){$size+=$package."Size (MB)"}
    "Total Size = $size MB"
    "Total Size = $($size/1024) GB"
}
else
{
    foreach($DP in $Unused)
    {
        set-location "$($SiteCode):"
        $DriverPackPath  = $DP.PkgSourcePath
        Remove-CMDriverPackage $DP
        #Cannot access the files from CM-Site. 
        #Set location back to an area that can access these files
        Set-Location C:
        Remove-Item -Path $DriverPackPath -Recurse -Force
    }
}

#Reset location
Set-Location $curLocation
