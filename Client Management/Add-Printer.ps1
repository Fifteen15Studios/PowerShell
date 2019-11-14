###############################################################################
#
# For this script to work properly, you may need the proper driver files to be
# in a subfolder called "drivers", and you will need to know what the name of
# the driver is within Windows. If you are using a built-in driver, the driver
# files will not be needed, but you still need to know the name of the driver
# within Windows.
#
# If the printer is already installed on a PC, the name of the driver can be
# found on the "Printer Properties" window, under the "Advanced" page.
#
###############################################################################

$IPAddress = #Example: "10.10.10.250"
$Driver = # Example: "HP Color LaserJet Pro M478f-9f PCL-6 (V4)"
$Name = # Example: "Front-Desk-LJ"
$INFFile = # Example: "hpclC62A4_x64.inf"

# Attempt to retrieve information about the printer
$Printer = Get-Printer -Name $Name -ErrorAction SilentlyContinue

# Check to see if printer exists
if($Printer -and ($Printer.PortName -eq $IPAddress))
{
    Write-Host "Printer already installed. No changes have been made." -ForegroundColor Green
    exit 0
}
# If Printer exists but with different IP, change IP
elseif($Printer)
{
    Write-Host "Printer Exists, but with the wrong IP. Changing IP address..." -ForegroundColor Yellow
    $OldIP = $Printer.PortName
    "Old IP: $OldIP"
    "New IP: $IPAddress"

    # Test connection to printer
    if(Test-Connection $IPAddress -Count 1)
    {
        # Add new IP as port
        if(-not (Get-PrinterPort -Name $IPAddress -ErrorAction SilentlyContinue))
        {
            Add-PrinterPort -Name $IPAddress -PrinterHostAddress $IPAddress
        }
        # Change IP
        Set-Printer -Name $Name -PortName $IPAddress
        # Remove old IP port
        Remove-PrinterPort -Name $OldIP -ErrorAction SilentlyContinue
    }
    else
    {
        Write-Error "Cannot connect to $IPAddress. Check your connection and try again. No changes have been made."
        exit 1
    }
}
# If printer has never been added
else
{
    "Printer does not exist. Adding printer..."

    # Test connection to printer
    if(Test-Connection $IPAddress -Count 1)
    {
        # Quit on error from here on out
        $ErrorActionPreference = "Stop"

        # If driver doesn't already exist, install it
        if(-not ($Driver -in (Get-PrinterDriver).Name))
        {
            # Get current path
            if ($psISE)
            {
                $ScriptPath = Split-Path -Parent $psISE.CurrentFile.FullPath
            }
            elseif($PSVersionTable.PSVersion.Major -gt 3)
            {
                $ScriptPath = $PSScriptRoot
            }
            else
            {
                $ScriptPath = split-path -parent $MyInvocation.MyCommand.Definition
            }

            # Add driver to Windows driver store
            pnputil.exe -add-driver "$ScriptPath\Driver\$INFFile"
            # Add printer driver
            Add-PrinterDriver $Driver
        }

        # Add new IP as port
        if(-not (Get-PrinterPort -Name $IPAddress -ErrorAction SilentlyContinue))
        {
            Add-PrinterPort -Name $IPAddress -PrinterHostAddress $IPAddress
        }
        # Add Printer
        Add-Printer -Name $Name -DriverName $Driver -PortName $IPAddress

        Write-Host "Successfully added $Name" -ForegroundColor Green
    }
    else
    {
        Write-Error "Cannot connect to printer. Check your connection and try again."
        Exit 1
    }
}
