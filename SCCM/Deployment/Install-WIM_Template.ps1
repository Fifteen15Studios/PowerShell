# This script is designed to deploy software that is packaged into a WIM file. To use it, create a WIM file with the required
# install files (including an install Script inside of the WIM) and then use this script to perform the install.
#
#
# To Create the WIM file, run the following command:
#
# New-WindowsImage -ImagePath "<path to save to>\<name>.wim" -CapturePath "<Path of source files>" -Name "InstallerSources"
#
#
# To Modify files inside of the WIM, first mount the WIM, then edit the file, then unmount the WIM with the save command:
#
# Mount-WindowsImage -ImagePath $WIMFile -Index 1 -Path $mountPath | Out-Null
# <Modify Files>
# Dismount-WindowsImage -Path $mountPath -Save
#
#
# PROPER USAGE FROM SCCM - "cmd /c start /wait powershell -executionpolicy bypass -file <Filename>.ps1"
# If you simply run "powershell -executionpolicy bypass -file <Filename>.ps1" it will run in the background without showing the PS window
#
# More info: https://adminsccm.com/2020/07/20/use-a-wim-to-deploy-large-apps-via-configmgr-app-model/

param(
    [switch]$Uninstall
)

#requires -runasadministrator

$Server = ""
$Share = ""
$AppName=""

$CSVPath = "\\$Server\$Share\Logs\$AppName\$AppName.csv"

$WIMFile = "$PSScriptRoot\<FileName>.wim"

# mount the WIM here
$mountPath = "$PSScriptRoot\mount"
$installScript = "$mountPath\<Install Script In WIM>.ps1"

#--------------------- DO NOT CHANGE ANYTHING BELOW THIS LINE -------------------------------
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

try {

    # Create folder if not already created
    if(-not (Test-Path $mountpath)) {
        New-Item -Path $mountPath -ItemType Directory
    }

    # If WIM isn't already mounted
    if(-not (Get-WindowsImage -Mounted | where {$_.ImagePath -like "*$WIMFile"})) {
        log "Mounting WIM"
        Mount-WindowsImage -ImagePath $WIMFile -Index 1 -Path $mountPath | Out-Null
        log "Wim mounted successfully!"
    }
    else {
        "WIM Mounted already. Moving on."
    }
}
catch {

    log "Failed to mount WIM"

    ## failed to mount the WIM, can't proceed
    exit 1
}

$argumentList = "-executionpolicy bypass -file `"$installScript`""

if($Uninstall) {
    $argumentList += " -uninstall"
}

## perform your install here - with return code
$returnCode = (Start-Process -Wait -FilePath powershell -ArgumentList $argumentList -PassThru).ExitCode

## dismount the WIM whether we succeeded or failed
try {
    log "Unmounting WIM."
    Get-WindowsImage -Mounted | ForEach-Object {$_ | Dismount-WindowsImage -Discard} | Out-Null
    #Dismount-WindowsImage -Path $mountPath -Discard
    log "WIM Unmounted successfully!"
}
catch {
    log "Unmounting failed. Scheduling unmount for next boot."

    ## failed to cleanly dismount, so set a task to cleanup after reboot
    $STAction = New-ScheduledTaskAction `
        -Execute 'Powershell.exe' `
        -Argument '-NoProfile -WindowStyle Hidden -command "& {Get-WindowsImage -Mounted | ForEach-Object {$_ | Dismount-WindowsImage -Discard -ErrorVariable wimerr; if ([bool]$wimerr) {$errflag = $true}}; If (-not $errflag) {Clear-WindowsCorruptMountPoint; Unregister-ScheduledTask -TaskName ''CleanupWIM'' -Confirm:$false}}"'
            
    $STTrigger = New-ScheduledTaskTrigger -AtStartup
        
    Register-ScheduledTask `
        -Action $STAction `
        -Trigger $STTrigger `
        -TaskName "CleanupWIM" `
        -Description "Clean up WIM Mount points that failed to dismount properly" `
        -User "NT AUTHORITY\SYSTEM" `
        -RunLevel Highest `
        -Force
}
    
## return exit code
exit $returnCode
