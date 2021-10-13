<#
# This is a template for installing and uninstalling software. It will log
# all installs in \\$Server\$Share\Logs\$AppName\$AppName.csv
# 
# Before performing the (un)install, it checks to see if a reboot is pending.
# If so, it will not continue on with the install and instead display a
# message to the user that they must first reboot.
# 
# Use the -Uninstall switch to perform an uninstall of the software. Otherwise
# it will perform an install.
#
# Required variables to set:
#   Server : Server name or IP Address to send logs to.
#   Share : Share name on the server.
#   AppName : Name of the application being (un)installed.
#   InstallFile : The file to use for the install. 
#     Example: "Install.exe" or "msiexec"
#   InstallArguments : Arguments to pass to the install file.
#     Example: "-i program.msi -qb"
#   UninstallFile : The file to use for the uninstall. 
#     Example: "Install.exe" or "msiexec"
#   UninstallArguments = Arguments to pass to the uninstall file.
#     Example: "-x program.msi -qn"
#
# PROPER USAGE FROM SCCM - "cmd /c start /wait powershell -executionpolicy bypass -file <Filename>.ps1"
# If you simply run "powershell -executionpolicy bypass -file <Filename>.ps1" it will run in the background without showing the PS window
#>


param(
    [switch]$Uninstall
)

$Server = ""
$Share = ""
$AppName=""

$CSVPath = "\\$Server\$Share\Logs\$AppName\$AppName.csv"

$InstallFile = "$PSScriptRoot\"
$InstallArguments = ""
$UninstallFile = "$PSScriptRoot\"
$UninstallArguments = ""

function Pre-Install() {
    # If you need to do anything before the install, put it here
}

function Post-Install() {
    # If you need to do anything after the install, put it here
}

function Pre-Uninstall() {
    # If you need to do anything before the uninstall, put it here
}

function Post-Uninstall() {
    # If you need to do anything after the uninstall, put it here
}

#--------------------- DO NOT CHANGE ANYTHING BELOW THIS LINE -------------------------------
# Tests if a reboot is pending
function Test-PendingReboot() {
    $PendingRebootTests = @(
        @{
            Name = 'RebootPending'
            Test = { Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing' -Name 'RebootPending' -ErrorAction Ignore }
            TestType = 'ValueExists'
        }
        @{
            Name = 'RebootRequired'
            Test = { Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update' -Name 'RebootRequired' -ErrorAction Ignore }
            TestType = 'ValueExists'
        }
    )

    foreach ($Test in $PendingRebootTests) {
        $Result = .($Test.Test)

        if ($Test.TestType -eq 'ValueExists' -and $Result) {
            return $Test.name
        } elseif ($Test.TestType -eq 'NonNullValue' -and $Result -and $Result.($Test.Name)) {
            return $Test.name
        }
    }

    return $False
}

# Writes $status to the csv file defined above.
# Includes a date/time stamp, and computer name.
function log($Status) {
    $Date = (Get-Date).ToString("yyyy/MM/dd")
    $Time = (Get-Date).ToString("HH:mm:ss")

    try {
        Add-Content -Value "$Date,$Time,$Env:Computername,$Status" -Path $CSVPath
    }
    catch {

    }
}

# If there is a reboot pending, don't do the install and display a message
if(Test-PendingReboot) {

    log "Reboot pending"

    Write-Host "A reboot is pending. Please reboot and try again." -BackgroundColor Red -ForegroundColor White
    
    $Seconds = 30
    $Length = $Seconds / 100
    $EndTime = [datetime]::UtcNow.AddSeconds($Seconds)

    while (($TimeRemaining = ($EndTime - [datetime]::UtcNow)) -gt 0) {
      Write-Progress -Activity "A reboot is pending. Please reboot and try again." -Status 'This window will close in 30 seconds.' -SecondsRemaining $TimeRemaining.TotalSeconds
      Start-Sleep 1
    }
    ""
    # 3010 reports that a reboot is necessary
    exit 3010
}

# If csv file doesn't exist, create it with the proper headings
# NOTE: This might throw an error if the folder isn't created or if the user doesn't have permissions
if(-not (Test-Path "$CSVPath")) {
    add-content -value 'Date,Time,Computername,Status' -path $CSVPath
}

# Warn the user not to close the window
Write-Host "DO NOT CLOSE THIS WINDOW!" -BackgroundColor Red -ForegroundColor White
""

# Perform the install or uninstall and get the exit code
if($Uninstall) {
    log "Uninstall_START"
    "Uninstalling $AppName. Please wait..."
    Pre-Install
    $ExitCode=(Start-Process -Wait -FilePath $UninstallFile -ArgumentList $UninstallArguments -PassThru).ExitCode
}
else {
    log "Install_START"
    "Installing $AppName. Please wait..."
    Pre-Uninstall
    $ExitCode=(Start-Process -Wait -FilePath $InstallFile -ArgumentList $InstallArguments -PassThru).ExitCode
}    

# Read exit code and report success or failure to the log before exiting with the same code
if($ExitCode -eq 0) {
    log "SUCCESS"
    
    If($uninstall) {
        Post-Uninstall
    }
    else {
        Post-Install
    }
}
else {
    log "Failed with exit code $ExitCode"
}

exit $ExitCode
