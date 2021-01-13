###############################################################################
#    
# This sript extracts the icon image from an exe file and saves it as an 
#   ico file. 
#
# When first run, you will be prompted to locate the exe file with an "Open File"
#   dialog. After the exe is selected, you will be prompted to choose a save 
#   location for the ico file. The icon from within the exe file will be 
#   extracted to the location, and file name, chosen in the "Save As" dialog.
#
###############################################################################

# Letter used for temporarily mapping a network drive when the exe is on a network drive
# Should be just a single letter, nothing more
$TempDriveLetter = "B"

# Show Open dialog, and return the path of the selected file
Function Get-FileName($initialDirectory)
{  
    # Create form object
    [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null

    # Create Open File Dialog
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    # Only accept .exe files
    $OpenFileDialog.filter = “EXE Files (*.exe)| *.exe”
    $OpenFileDialog.Title = "Select an EXE file"
    $OpenFileDialog.ShowDialog() | Out-Null

    # Return the value
    $OpenFileDialog.filename
}

# Show Save dialog and return path of file
Function Save-FileName($initialDirectory, $initialFileName) {

    # Create Save File Dialog
    $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $SaveFileDialog.initialDirectory = $initialDirectory
    # Only accept .ico files
    $SaveFileDialog.filter = “ICO Files (*.ico)| *.ico”
    # Set initial file name
    $SaveFileDialog.FileName = $initialFileName
    $SaveFileDialog.Title = "Save Icon As"
    $SaveFileDialog.ShowDialog() | Out-Null

    # Return the value
    $SaveFileDialog.filename
}

# Get the EXE you want to extract the icon from
$file = Get-FileName -initialDirectory $env:USERPROFILE

# Load System.Drawing library
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing')  | Out-Null

# If the file exists, which it should
if(Test-Path $file) {
    
    # Get the name of the file
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file)

    # Get the location to save to
    $saveName = Save-FileName $([System.IO.Path]::GetDirectoryName($file)) $baseName

    # If it's a network file, map a temporary network drive
    if($file -like "\\*") {
        New-SmbMapping -LocalPath "$($TempDriveLetter):" -RemotePath $([System.IO.Path]::GetDirectoryName($file)) | Out-Null
        $file = "$($TempDriveLetter):\$([System.IO.Path]::GetFileName($file))"
        $network = $true
    }

    # Extract the icon
    [System.Drawing.Icon]::ExtractAssociatedIcon($file).ToBitmap().Save("$saveName")

    # If file was on the network, remove network drive
    if($network) {
        Remove-SmbMapping "$($TempDriveLetter):" -Force
        Remove-Variable network
    }
}