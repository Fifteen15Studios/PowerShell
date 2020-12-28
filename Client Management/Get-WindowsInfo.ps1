###############################################################################
#
# This script gets a bunch of information about the current install of Windows
#
# This should work in any environment, as it does not require 
#    PowerShell remoting or any other specific environment setup to reach the
#    remote PC. It simply uses WMI to do all of its work on remote machines.
#
###############################################################################

$ComputerName = Read-Host "Enter a computer name (Leave blank for local machine)"

# Get Information
function Get-Info($ComputerName) {

    # If remote computer
    if($ComputerName -ne $env:COMPUTERNAME){     

        # Get-CimInstance doesn't work remotely for some reason - unless firewall is open and other items are in place
        # It also formats LastBootTime differently than Get-WmiObject
        $OS = Get-WmiObject -ComputerName $ComputerName Win32_OperatingSystem

        $Version = [Version]$OS.version

        # Registry variables - needed because of WMI remote registry method
        #$HKEY_CURRENT_USER = 2147483649
        $HKEY_Local_Machine =2147483650
        $Reg = [WMIClass]"\\$ComputerName\ROOT\DEFAULT:StdRegProv"
        $Key = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"

        # Values we need to get
        $ServerValue = "InstallationType"
        $BuildValue = "CurrentBuild"
        $InstallDateValue = "InstallDate"
        $ReleaseIDValue = "ReleaseId"

        # Read from remote registry, then get results from the read
        $Results = $Reg.GetStringValue($HKEY_Local_Machine, $Key, $ServerValue)
        $ServerOS = $results.sValue -eq "Server"

        $Results = $Reg.GetStringValue($HKEY_Local_Machine, $Key, $BuildValue)
        $Build = $results.sValue

        # Format the install date properly
        $InstallDate = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970')).AddSeconds($($Reg.GetDWORDValue($HKEY_Local_Machine, $Key, $InstallDateValue).uValue))
        
        # If Windows 10, get ReleaseID
        if($Version.Major -ge 10){
            $Results = $Reg.GetStringValue($HKEY_Local_Machine, $Key, $ReleaseIDValue)
            $ReleaseID = $Results.sValue
        }
    }
    # If local computer
    else {
        # Do not change from Get-WmiObject to Get-CimInstance. LastBootTime is formatted diferently in the 2 functions.
        $OS = Get-WmiObject Win32_OperatingSystem

        $Version = [Version]$OS.version

        # Get values from registry
        $Key = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
        $ServerOS = (Get-ItemProperty $Key).InstallationType -eq "Server"
        $Build = (Get-ItemProperty $Key).CurrentBuild
        $InstallDate = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970')).AddSeconds($(Get-ItemProperty $Key).InstallDate)

        # If Windows 10, get Release ID
        if($Version.Major -ge 10){
            $ReleaseID=(Get-ItemProperty $Key).ReleaseId
        }
    }

    # These are retrieved the same way whether local or remote
    $Name = $OS.caption
    $Bit = $OS.OSArchitecture
    $LastBootUpTime = $OS.ConvertToDateTime($OS.LastBootUpTime)

    # Put Information into an object
    $Info = New-Object -TypeName psobject

    $Info | Add-Member -MemberType NoteProperty -Name ComputerName -Value $OS.PSComputerName
    $Info | Add-Member -MemberType NoteProperty -Name OS_Name -Value $Name
    $Info | Add-Member -MemberType NoteProperty -Name Bits -Value $Bit
    $Info | Add-Member -MemberType NoteProperty -Name Server -Value $ServerOS
    $Info | Add-Member -MemberType NoteProperty -Name Version -Value $Version
    # If Windows 10, Add Release ID
    if($Version.Major -ge 10){$Info | Add-Member -MemberType NoteProperty -Name ReleaseID -Value $ReleaseID}
    $Info | Add-Member -MemberType NoteProperty -Name Build -Value $Build
    $Info | Add-Member -MemberType NoteProperty -Name InstallDate -Value $InstallDate

    # Calculate uptime based on current time and last boot time
    $uptime = ((get-date) - $LastBootUpTime)


    # Format uptime for display
    $Output = ""

    if($uptime.Days -gt 1)
    {$Output += $uptime.Days.ToString() + " Days "}
    elseif($uptime.Days -eq 1)
    {$Output += "1 Day "}

    if($uptime.Hours -gt 1)
    {$Output += $uptime.Hours.ToString() + " Hours "}
    elseif($uptime.Hours -eq 1)
    {$Output += "1 Hour "}

    if($uptime.Minutes -gt 1) 
    {$Output += $uptime.Minutes.ToString() + " Minutes "}
    elseif($uptime.Minutes -eq 1) 
    {$Output += "1 Minute "}

    if($uptime.Seconds -gt 1)
    {$Output += $uptime.Seconds.ToString() + " Seconds "}
    elseif($uptime.Seconds -eq 1)
    {$Output += "1 Second "}

    $Info | Add-Member -MemberType NoteProperty -Name UpTime -Value ($($Output.trim()))

    # Display information
    $Info
}

# If computername is empty, use local computer
if(-not [bool]$ComputerName) {
    $ComputerName = $env:COMPUTERNAME
}

Get-Info($ComputerName)
