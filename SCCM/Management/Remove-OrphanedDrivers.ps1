param(
    [Switch]
    $whatif
)

# Import SCCM module and set location to SCCM site
Import-Module "C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
$SiteCode = Get-PSDrive -PSProvider CMSITE

$oldLocation = Get-Location

# Location must be the SCCM drive for the next step
set-location "$($SiteCode):"

# Get list of drivers and driver packs from SCCM
"Getting drivers..."
""
#Suppress warning
$CMPSSuppressFastNotUsedCheck = $true
$Drivers = Get-CMDriver

# Create lists to hold items
$OrphanedDrivers = [System.Collections.ArrayList]@()

# Location must not be the SCCM drive for the next step
Set-Location C:

#for each driver, check to see if it's in a driver pack
foreach($Driver in $Drivers) {
    # If driver is not null, and it's path is not valid
    if($Driver -and (-not (test-path $Driver.ContentSourcePath))) {
        $OrphanedDrivers.add($Driver) | Out-Null
        #"$($Driver.LocalizedDisplayName) filepath not found"
    }
}

if(-not $whatif) {
    "Removing $($OrphanedDrivers.Count) orphaned drivers..."

    $CMPSSuppressFastNotUsedCheck = $false

    # Location must be the SCCM drive for the next step
    set-location "$($SiteCode):"

    foreach($driver in $OrphanedDrivers) {
        try{
            $driver | Remove-CMDriver -Force
        }
        catch {
            "Error removing `"$($driver.LocalizedDisplayName)`" at `"$($driver.ContentSourcePath)`"."
        }
    }
}
else {
    "$($OrphanedDrivers.Count) orphaned drivers exist."
}

Set-Location $oldLocation