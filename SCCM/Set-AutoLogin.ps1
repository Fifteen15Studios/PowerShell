#requires -runasadministrator
#
# Sets or Removes autologin capability on a Windows PC.
#
# SCCM scripting does not support switches, so I instead used an int for the 
# enable switch. If $enable is 0 then it will disable any existing auto logon,
# otherwise it will enable it using the username, password, and domain name 
# provided.
#
# If this is being used outside of SCCM and you would prefer to use a switch:
# 1) the [int] would be changed to [switch], 
# 2) the HelpMessage could be changed or removed, and remove mandatory
# 3) the [ValidateRange(0,1)] would nee to be removed, 
# 4)  "$enable -eq 1" would be changed to "$enable".
# Or, $Enable can be changed to $Disable in the parameters, with the changes
# listed above, and in step 4 "$enable -eq 1" would become "-not $Disable"

param(
        [parameter(HelpMessage="Auto Login username")]
        [string]$UserName,
        [parameter(HelpMessage="Auto Login password")]
        [string]$Password,
        [parameter(HelpMessage="Login Domain")]
        [string]$Domain,
        [parameter(mandatory=$true,HelpMessage="1 for enable, 0 for disable")]
        [ValidateRange(0,1)]
        [int]$Enable
)

#Registry path and keys
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon"
$AutoLoginKey = "AutoAdminLogon"
$UsernameKey = "DefaultUserName"
$PasswordKey = "DefaultPassword"
$DomainKey = "DefaultDomainName"

New-ItemProperty -Path $RegPath -Name $AutoLoginKey -Value $Enable -Force

if($Enable -eq 1)
{
    New-ItemProperty -Path $RegPath -Name $UsernameKey -Value $Username -Force
    New-ItemProperty -Path $RegPath -Name $PasswordKey -Value $Password -Force
    New-ItemProperty -Path $RegPath -Name $DomainKey -Value $Domain -Force
}
else
{
    Remove-ItemProperty -Path $RegPath -Name $UsernameKey -Force
    Remove-ItemProperty -Path $RegPath -Name $PasswordKey -Force
    Remove-ItemProperty -Path $RegPath -Name $DomainKey -Force
}
