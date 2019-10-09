# Purpose: 
# Find the uninstall string for a specific application.
# This can then be used to perform an uninstall.
#
# Parameters: 
# appName - Name of 1 or more applications to find the Uninstall string for

param( 
  [parameter(ValueFromPipeline=$true, Mandatory=$true, HelpMessage="App you would like to remove")]
  [string[]]$appName
)

$RegPath32 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
$RegPath64 = "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"

#Seach 64-bit environment
if([environment]::Is64BitOperatingSystem)
{
    foreach($app in $appname)
    {Get-ChildItem -Path $RegPath64 | 
        Get-ItemProperty |
        Where-Object {$_.DisplayName -like "*$app*" } |
        Select-Object -Property DisplayName, UninstallString}
}

#Search for 32-bit apps regardless of OS architecture
foreach($app in $appname)
    {Get-ChildItem -Path $RegPath32 | 
        Get-ItemProperty |
        Where-Object {$_.DisplayName -like "*$app*" } |
        Select-Object -Property DisplayName, UninstallString}
