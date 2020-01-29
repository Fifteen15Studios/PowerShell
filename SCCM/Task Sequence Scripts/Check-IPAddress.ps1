# This can be used in a task sequence to check to see if there is an IP address 
# assigned to the adapter. If there isn't an IP, display a message to the user
# asking them to disconnect and reconnect the cable, wait 10 seconds, 
# then check again for an IP in a while loop.

#Hide the progress dialog
$TSProgressUI = new-object -comobject Microsoft.SMS.TSProgressUI
$TSProgressUI.CloseProgressDialog()

# Hard-coded variables
$title = "No IP address"
$message = "Network adapter was not activated after driver install.

Remove the adapter, reinsert it,
and press Enter [OK]."

<#You can use task sequence variables for these values by uncommenting 
# this block of code

# connect to Task Sequence environment
$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
# read variables set in the task sequence
$title = $tsenv.Value("title")
$message = $tsenv.Value("message")#>

$result = gwmi -Query 'SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled="True"'

while($result -eq $null)
{
    # now show a popup message to the end user
    write-host "`n`n$title`n`n$message"
    $form = [System.Reflection.Assembly]::LoadWithPartialName(“System.Windows.Forms”)
    $result = [Windows.Forms.MessageBox]::Show(“$message”, “$title”, [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)
    Start-Sleep -s 10
    
    $result = gwmi -Query 'SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled="True"'
}
