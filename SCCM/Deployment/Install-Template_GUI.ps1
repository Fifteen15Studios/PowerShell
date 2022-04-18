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
#   Server             : Server name or IP Address to send logs to.
#   Share              : Share name on the server to send logs to.
#   AppName            : Name of the application.
#   InstallFile        : The file to use for the install. In most cases, should start with $PSScriptRoot\
#     Example            : "$PSScriptRoot\Install.exe" or "msiexec"
#   InstallArguments   : Arguments to pass to the install file.
#     Example            : "--silent" or "-i program.msi -qb"
#
# Optional variables to set: 
# If you want to use uninstall, or reinstall switches, uninstall values are required:
#   UninstallFile      : The file to use for the uninstall. In most cases, should start with $PSScriptRoot\
#   UninstallArguments : Arguments to pass to the uninstall file.
# If you want to use repair reinstall switches, repair values are required:
#   RepairFile         : The file to use for the repair. In most cases, should start with $PSScriptRoot\
#   RepairArguments    : Arguments to pass to the repair file.
#
# Switches:
#   Uninstall : Perform an uninstall of the software using UninstallFile and UninstallArguments
#   Reinstall : Perform an uninstall then install of the software
#   Repair    : Perform a repair of the software using RepairFile and RepairArguments
#
# PROPER USAGE FROM SCCM - 
# "cmd /c start /wait powershell -executionpolicy bypass -file <Filename>.ps1" To show PowerShell window
# "powershell -executionpolicy bypass -file <Filename>.ps1" to run in the background without showing the PowerShell window
#>

param(
    [switch]$Uninstall,
    [switch]$Reinstall,
    [switch]$Repair
)

# ----------------------------- FILL IN THIS LINE -----------------------------
$AppName=""

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework # For MessageBox

$Script = [scriptblock]::Create({ 
param(
    $params
)

$Uninstall = $params.Uninstall
$Reinstall = $params.Reinstall
$Repair = $params.Repair
$AppName = $params.AppName
$ScriptRoot = $params.ScriptRoot

Set-Location $ScriptRoot

# Codes other than 0 means reboot required, but was successful
$SuccessCodes = (0, 1641, 3010, 17022)

# ------------------------- CHANGE THESE VARIABLES ----------------------------
$Server = "VRThor2012"
$Share = "Software"

$CSVPath = "\\$Server\$Share\Logs\$AppName\$AppName.csv"

$InstallFile = "$ScriptRoot\"
$InstallArguments = ""
$UninstallFile = "$ScriptRoot\"
$UninstallArguments = ""
$RepairFile = "$ScriptRoot\"
$RepairArguments = ""

Write-Host "installFile = $InstallFile"

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

    # Create folder structure
    if(-not (Test-Path (Split-Path $CSVPath))) {
        New-Item (Split-Path $CSVPath) -ItemType Directory | Out-Null
    }

    # If csv file doesn't exist, create it with the proper headings
    # NOTE: This might throw an error if the folder isn't created or if the user doesn't have permissions
    if(-not (Test-Path "$CSVPath")) {
        Add-Content -Value 'Date,Time,Computername,Status' -Path $CSVPath
    }

    $Date = (Get-Date).ToString("yyyy/MM/dd")
    $Time = (Get-Date).ToString("HH:mm:ss")

    try {
        Add-Content -Value "$Date,$Time,$Env:Computername,$Status" -Path $CSVPath
    }
    catch {

    }
}

# Perform the install or uninstall and get the exit code
function Uninstall() {
    log "Uninstall_START"
    "Uninstalling $AppName. Please wait..."
    Pre-Uninstall
    if($UninstallArguments) {
        $script:ExitCode=(Start-Process -Wait -FilePath $UninstallFile -ArgumentList $UninstallArguments -PassThru -WindowStyle Hidden).ExitCode
    }
    else {
        $script:ExitCode=(Start-Process -Wait -FilePath $UninstallFile -PassThru -WindowStyle Hidden).ExitCode
    }
}
function Install() {
    log "Install_START"
    "Installing $AppName. Please wait..."
    Pre-Install
    if($InstallArguments) {
        $script:ExitCode=(Start-Process -Wait -FilePath $InstallFile -ArgumentList $InstallArguments -PassThru -WindowStyle Hidden).ExitCode
    }
    else {
        $script:ExitCode=(Start-Process -Wait -FilePath $InstallFile -PassThru -WindowStyle Hidden).ExitCode
    }
    $script:install = $true
}
function Repair() {
    log "Repair_START"
    "Repairing $AppName. Please wait..."
    if($RepairArguments) {
        $script:ExitCode=(Start-Process -Wait -FilePath $RepairFile -ArgumentList $RepairArguments -PassThru -WindowStyle Hidden).ExitCode
    }
    else {
        $script:ExitCode=(Start-Process -Wait -FilePath $RepairFile -PassThru -WindowStyle Hidden).ExitCode
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

# Warn the user not to close the window
Write-Host "DO NOT CLOSE THIS WINDOW!" -BackgroundColor Red -ForegroundColor White
""

# Perform the appropriate action
if($Repair) {
    Repair
}
elseif($Reinstall) {
    Uninstall
    # If uninstall was successful
    if($ExitCode -in $SuccessCodes) {
        Post-Uninstall
        Install
    }
}
elseif($Uninstall) { Uninstall }
else { Install }

# Read exit code and report success or failure to the log before exiting with the same code
if($ExitCode -in $SuccessCodes) {
    log "SUCCESS"
    
    If($uninstall) {
        Post-Uninstall
    }
    # If install or reinstall
    elseif($Install) {
        Post-Install
    }
}
else {
    log "Failed with exit code $ExitCode"
}

exit $ExitCode
})

function Show-Form($title, $scriptBlock) {
    $script:main_form = New-Object System.Windows.Forms.Form
    $main_form.Text = $title

    $Label1 = New-Object System.Windows.Forms.Label
    $Label1.Text = "$title. This window will close automatically when finished.`r`n`r`nThis may take a while. Please wait..."
    $Label1.Location  = New-Object System.Drawing.Point(10,10)
    $Label1.AutoSize = $true

    $main_form.controls.Add($Label1)

    $Label2 = New-Object System.Windows.Forms.Label
    $Label2.Text = "This window will close on its own when finished.`r`nPlease do not close this window unless directed to do so by the IT or Design Technology team."
    $Label2.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 14, [System.Drawing.FontStyle]::Bold)
    $Label2.Location  = New-Object System.Drawing.Point(10, ($Label1.Bottom + 10))
    $Label2.AutoSize = $true

    $main_form.controls.Add($Label2)

    $main_form.AutoSize = $true

    # Make it open center screen
    $main_form.StartPosition = 1
    # Make form not resizable
    $main_form.FormBorderStyle = 'Fixed3D'
    $main_form.MaximizeBox = $false
    $main_form.MinimizeBox = $false
    $main_form.ControlBox = $False

    $main_form.ShowIcon = $false

    $main_form.Add_Shown({
            # Bring the form to the front
            $this.Activate()
            $this.topMost = $true
            $this.topMost = $false

            write-host "AppName before: $AppName"
            $myArgs = @{
                Uninstall = $Uninstall
                Reinstall = $Reinstall
                Repair = $Repair
                AppName = $AppName
                ScriptRoot = $PSScriptRoot
            }
            
            # Run the install in the background - allowing the form to move
            $job = start-job -scriptBlock $scriptBlock -ArgumentList $myArgs

            # Wait until it is complete
            Do {[System.Windows.Forms.Application]::DoEvents()} Until ($job.State -eq "Completed")

            $exitCode = $job | Receive-Job

            $main_form.close()
            $LASTEXITCODE = $exitCode
    })

    $main_form.AutoScroll = $true

    $main_form.ShowDialog() | Out-Null
}

If($Uninstall) {
    $title = "Uninstalling $AppName"
}
elseIf($reinstall) {
    $title = "Reinstalling $AppName"
}
elseIf($repair) {
    $title = "Repairing $AppName"
}
else {
    $title = "installing $AppName"
}

Show-Form $title $Script