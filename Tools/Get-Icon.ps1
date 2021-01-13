###############################################################################
#    
# This sript extracts the icon image from an EXE file and saves it as an 
#   ICO file. 
#
# When first run, you will be prompted to locate the EXE file with an "Open File"
#   dialog. After the exe is selected, you will be prompted to choose a save 
#   location for the ICO file. The icon from within the exe file will be 
#   extracted to the location, and file name, chosen in the "Save As" dialog.
#
###############################################################################

# Letter used for temporarily mapping a network drive when the EXE is on a network drive
# Should be just a single letter, nothing more
$TempDriveLetter = "B"

# Show Open dialog, and return the path of the selected file
Function Get-FileName($InitialDirectory)
{  
    # Create form object
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null

    # Create Open File Dialog
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    # Only accept .exe files
    $OpenFileDialog.Filter = “EXE Files (*.exe)| *.exe”
    $OpenFileDialog.Title = "Select an EXE file"
    $Result = $OpenFileDialog.ShowDialog()

    if($Result -eq "OK") {
        # Return the value
        $OpenFileDialog.FileName
    }
    # return an empty string if cancelled
    else {
        ""
    }
}

# Show Save dialog and return path of file
Function Save-FileName($InitialDirectory, $InitialFileName) {

    # Create Save File Dialog
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.InitialDirectory = $InitialDirectory
    # Only accept .ico files
    $SaveFileDialog.Filter = “ICO Files (*.ico)| *.ico”
    # Set initial file name
    $SaveFileDialog.FileName = $InitialFileName
    $SaveFileDialog.Title = "Save Icon As"
    $Result = $SaveFileDialog.ShowDialog()

    if($Result -eq "OK") {
        # Return the value
        $SaveFileDialog.FileName
    }
    # return an empty string if cancelled
    else {
        ""
    }
}

# Get the EXE you want to extract the icon from
$File = Get-FileName -initialDirectory $env:USERPROFILE

# If dialog was cancelled, break out of the script
if(!$file) {
    break script
}

# Load System.Drawing library
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')  | Out-Null

# If the file exists, which it should
if(Test-Path $File) {
    
    # Get the name and directory of the file
    $BaseDir = [System.IO.Path]::GetDirectoryName($File)
    $BaseName = [System.IO.Path]::GetFileNameWithoutExtension($File)

    # Get the location to save to
    $SaveName = Save-FileName $BaseDir $BaseName

    # If dialog was cancelled, break out of the script
    if(!$SaveName) {
        break script
    }

    # If it's a network file, map a temporary network drive
    if($File -like "\\*") {
        New-SmbMapping -LocalPath "$($TempDriveLetter):" -RemotePath $BaseDir | Out-Null
        $File = "$($TempDriveLetter):\$([System.IO.Path]::GetFileName($File))"
        $Network = $true
    }

    # Extract the icon
    [System.Drawing.Icon]::ExtractAssociatedIcon($File).ToBitmap().Save("$SaveName")

    # If file was on the network, remove network drive
    if($Network) {
        Remove-SmbMapping "$($TempDriveLetter):" -Force
        Remove-Variable network
    }
}
