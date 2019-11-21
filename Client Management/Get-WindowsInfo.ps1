###############################################################################
#
# This script gets a bunch of information about the current install of Windows
#
###############################################################################

# Get Information
$Name = (Get-CimInstance Win32_OperatingSystem).caption
$ServerOS = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").InstallationType -eq "Server"
$Bit = (Get-CimInstance Win32_OperatingSystem).OSArchitecture
$Version = [Version](Get-CimInstance win32_operatingsystem).version
$Build= (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
$InstallDate = [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970')).AddSeconds($(Get-ItemProperty "HKLM:\Software\Microsoft\Windows NT\CurrentVersion").InstallDate)

# If Windows 10, get Release ID
if($Version.Major -ge 10){$ReleaseID=(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ReleaseId}

function getUptime
{
    $uptime = ((get-date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime)

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

    "$($Output.trim())"
}

# Put Information into an object
$Info = New-Object -TypeName psobject

$Info | Add-Member -MemberType NoteProperty -Name Name -Value $Name
$Info | Add-Member -MemberType NoteProperty -Name Bits -Value $Bit
$Info | Add-Member -MemberType NoteProperty -Name Server -Value $ServerOS
$Info | Add-Member -MemberType NoteProperty -Name Version -Value $Version
# If Windows 10, Add Release ID
if($Version.Major -ge 10){$Info | Add-Member -MemberType NoteProperty -Name ReleaseID -Value $ReleaseID}
$Info | Add-Member -MemberType NoteProperty -Name Build -Value $Build
$Info | Add-Member -MemberType NoteProperty -Name InstallDate -Value $InstallDate
$Info | Add-Member -MemberType NoteProperty -Name UpTime -Value (getUptime)

# Display information
$Info
