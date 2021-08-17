<#
.SYNOPSIS
    Install RSAT features for Windows 10 1809+
    
.DESCRIPTION
    Install RSAT features for Windows 10 1809+. All features are installed online from Microsoft Update thus the script requires Internet access

.PARAM InstallType
    What parts of RSAT to install. Accepts 3 possible values:
        All - Installs all the features within RSAT. This takes several minutes, depending on your Internet connection
        Basic - Installs ADDS, DHCP, DNS, GPO, ServerManager (If no parameters are specified, this is default behavior.)
        ServerManager - Installs ServerManager

.PARAM Uninstall
    Uninstalls all the RSAT features

.NOTES
    Filename: Install-RSAT.ps1
    Version: 1.3RH
    Author: Martin Bengtsson
    Blog: www.imab.dk
    Twitter: @mwbengtsson

    Version history:

    1.0   -   Script created

    1.2   -   Added test for pending reboots. If reboot is pending, RSAT features might not install successfully
              Added test for configuration of local WSUS by Group Policy.
                - If local WSUS is configured by Group Policy, history shows that additional settings might be needed for some environments
   
    1.3RH -   Changed parameters to a single parameter for InstallType
              Changed version check to simple check if it's Windows 10 1809 or higher
              Fixed some typos
              Added runasadmin in requires statement instead of function check
              Added check for both parameters, and throw error if both present
              Fixed some broken logic 
                - checking '-eq "True"'
                - outputing $RsatItem when it hadn't been set

    1.4RH -   If WSUS Server setting is found, attempts to disable and the re-enable the setting in the registry
                - This may not work if it's set via Group Policy
#>

#Requires -runasadmin

param(
    [parameter(Mandatory=$false)]
    [ValidateSet("All","Basic","ServerManager")]
    [string]$InstallType,
    [parameter(Mandatory=$false)]
    [ValidateNotNullOrEmpty()]
    [switch]$Uninstall
)

# Create Pending Reboot function for registry
function Test-PendingRebootRegistry {
    $CBSRebootKey = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction Ignore
    $WURebootKey = Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction Ignore
    
    if (($CBSRebootKey) -OR ($WURebootKey)) {
        $true
    }
    else {
        $false
    }
}

if($PSBoundParameters["InstallType"] -and $PSBoundParameters["Uninstall"])
{
    Write-Error "Cannot use both parameters. If you would like to uninstall, please only use -Uninstall"
    exit 1
}

# Windows 10 1809 build
$1809Build = "17763"

# Check if running Windows 10
if(([Version](Get-CimInstance win32_operatingsystem).version).Major -ge 10)
{
    # Get running Windows build
    $WindowsBuild = (Get-WmiObject -Class Win32_OperatingSystem).BuildNumber
}
else
{
    Write-Host "Need Windows 10 or higher." -ForegroundColor Red
    exit 1
}

# Get information about local WSUS server
$WUServer = (Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name UseWUServer -ErrorAction Ignore).UseWUServer -eq 1
$WSUSKeySet = $false
$TestPendingRebootRegistry = Test-PendingRebootRegistry

if (($WindowsBuild -ge $1809Build)) {
    Write-Verbose -Verbose "Running correct Windows 10 build number for installing RSAT with Features on Demand. Build number is: $WindowsBuild"
    Write-Verbose -Verbose "***********************************************************"

    if ($WUServer) {
        Write-Verbose -Verbose "A local WSUS server was found configured by group policy: $WUServer"
        Write-Verbose -Verbose "You might need to configure additional setting by GPO if things are not working"
        Write-Verbose -Verbose "The GPO of interest is following: Specify settings for optional component installation and component repair"
        Write-Verbose -Verbose "Check ON: Download repair content and optional features directly from Windows Update..."
        Write-Verbose -Verbose "Attempting to disable WSUS Setting..."
        Set-ItemProperty -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "UseWUServer" -Value 0 | Out-Null
        # If successfully set key
        if($?) {
            Write-Verbose -Verbose "Success!"
            $WSUSKeySet = $true
        }
        else {
            Write-Verbose -Verbose "Failed to write registry key. GPO may be active."
        }

        Write-Verbose -Verbose "***********************************************************"
    }

    if ($TestPendingRebootRegistry) {
        Write-Verbose -Verbose "Reboots are pending. The script will continue, but RSAT might not install successfully"
        Write-Verbose -Verbose "***********************************************************"
    }

    if ($PSBoundParameters["InstallType"] -eq "All") {
        Write-Verbose -Verbose "Script is running with -InstallType All. Installing all available RSAT features"
        $Install = Get-WindowsCapability -Online | Where-Object {$_.Name -like "Rsat*" -AND $_.State -eq "NotPresent"}
        if ($Install -ne $null) {
            foreach ($Item in $Install) {
                $RsatItem = $Item.Name
                Write-Verbose -Verbose "Adding $RsatItem to Windows"
                try {
                    Add-WindowsCapability -Online -Name $RsatItem
                }
                catch [System.Exception]
                {
                    Write-Verbose -Verbose "Failed to add $RsatItem to Windows"
                    Write-Warning -Message $_.Exception.Message
                }
            }
        }
        else {
            Write-Verbose -Verbose "All RSAT features seems to be installed already"
        }
    }

    # Assume Basic if no parameters provided
    if ($PSBoundParameters["InstallType"] -eq "Basic" -or $PSBoundParameters.Count -eq 0) {
        Write-Verbose -Verbose "Script is running with -InstallType Basic. Installing basic RSAT features"
        # Querying for what I see as the basic features of RSAT. Modify this if you think something is missing. :-)
        $Install = Get-WindowsCapability -Online | Where-Object {$_.Name -like "Rsat.ActiveDirectory*" -OR $_.Name -like "Rsat.DHCP.Tools*" -OR $_.Name -like "Rsat.Dns.Tools*" -OR $_.Name -like "Rsat.GroupPolicy*" -OR $_.Name -like "Rsat.ServerManager*" -OR $_.name -like "RSAT.BitLocker*" -AND $_.State -eq "NotPresent" }
        if ($Install -ne $null) {
            foreach ($Item in $Install) {
                $RsatItem = $Item.Name
                Write-Verbose -Verbose "Adding $RsatItem to Windows"
                try {
                    Add-WindowsCapability -Online -Name $RsatItem
                }
                catch [System.Exception]
                {
                    Write-Verbose -Verbose "Failed to add $RsatItem to Windows"
                    Write-Warning -Message $_.Exception.Message
                }
            }
        }
        else {
            Write-Verbose -Verbose "The basic features of RSAT seems to be installed already"
        }
    }

    if ($PSBoundParameters["InstallType"] -eq "ServerManager") {
        Write-Verbose -Verbose "Script is running with -InstallType ServerManager. Installing Server Manager RSAT feature"
        $Install = Get-WindowsCapability -Online | Where-Object {$_.Name -like "Rsat.ServerManager*" -AND $_.State -eq "NotPresent"} 
        if ($Install -ne $null) {
            $RsatItem = $Install.Name
            Write-Verbose -Verbose "Adding $RsatItem to Windows"
            try {
                Add-WindowsCapability -Online -Name $RsatItem
            }
            catch [System.Exception]
            {
                Write-Verbose -Verbose "Failed to add $RsatItem to Windows"
                Write-Warning -Message $_.Exception.Message ; break
            }
        }
        
        else {
            Write-Verbose -Verbose "ServerManager seems to be installed already"
        }
    }

    if ($PSBoundParameters["Uninstall"]) {
        Write-Verbose -Verbose "Script is running with -Uninstall parameter. Uninstalling all RSAT features"
        # Querying for installed RSAT features first time
        $Installed = Get-WindowsCapability -Online | Where-Object {$_.Name -like "Rsat*" -AND $_.State -eq "Installed" -AND $_.Name -notlike "Rsat.ServerManager*" -AND $_.Name -notlike "Rsat.GroupPolicy*" -AND $_.Name -notlike "Rsat.ActiveDirectory*"} 
        if ($Installed -ne $null) {
            Write-Verbose -Verbose "Uninstalling the first round of RSAT features"
            # Uninstalling first round of RSAT features - some features seems to be locked until others are uninstalled first
            foreach ($Item in $Installed) {
                $RsatItem = $Item.Name
                Write-Verbose -Verbose "Uninstalling $RsatItem from Windows"
                try {
                    Remove-WindowsCapability -Name $RsatItem -Online
                }
                catch [System.Exception]
                {
                    Write-Verbose -Verbose "Failed to uninstall $RsatItem from Windows"
                    Write-Warning -Message $_.Exception.Message
                }
            }       
        }
        # Querying for installed RSAT features second time
        $Installed = Get-WindowsCapability -Online | Where-Object {$_.Name -like "Rsat*" -AND $_.State -eq "Installed"}
        if ($Installed -ne $null) { 
            Write-Verbose -Verbose "Uninstalling the second round of RSAT features"
            # Uninstalling second round of RSAT features
            foreach ($Item in $Installed) {
                $RsatItem = $Item.Name
                Write-Verbose -Verbose "Uninstalling $RsatItem from Windows"
                try {
                    Remove-WindowsCapability -Name $RsatItem -Online
                }
                catch [System.Exception]
                {
                    Write-Verbose -Verbose "Failed to remove $RsatItem from Windows"
                    Write-Warning -Message $_.Exception.Message
                }
            } 
        }
        else {
            Write-Verbose -Verbose "All RSAT features seems to be uninstalled already"
        }
    }

    # If WSUS Setting was changed in registry, change it back
    if($WSUSKeySet) {
        Write-Verbose -Verbose "Attempting to re-enable WSUS Setting..."
            Set-ItemProperty -path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "UseWUServer" -Value 1 | Out-Null
            # If successfully set key
            if($?) {
                Write-Verbose -Verbose "Success!"
            }
            else {
                Write-Verbose -Verbose "Failed to set HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\UseWUServer to its original value (1)"
            }
    }
}
else {
    Write-Warning -Message "Must be running Windows 10 build $1809Build or higher. Your build: $WindowsBuild"
}
