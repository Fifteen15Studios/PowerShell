#Requires -runasadministrator
#
# Changes the password of a local user on a Windows machine. Also can enable or
# disable the account by using the appropriate switch.
#
# Parameters: 
#    Username - The username whose password you want to change
#    Password - The new password
#    enable (Optional) - if set, forces the user to be enabled on the machine
#    disable (Optional) - If set, forces the user to be disabled on the machine


param (
    [parameter(mandatory=$true)][string]$username,
    [parameter(mandatory=$true)][Security.SecureString]$password,
    [switch]$enable,
    [switch]$disable)

if($enable -and $disable)
{
    Write-Warning "Cannot use both ""Enable"" and ""Disable"" switches."
    exit
}

function setPassword($Password)
{  
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($password)
    $plainTextPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    try{
        $ADSIUser.setPassword("$plainTextPassword")
        $ADSIUser.setInfo()
        "Changed $username's password"
    }
    catch{
        write-host $error[0] -ForegroundColor Red
        exit
    }
    finally
    {
        Remove-Variable plainTextPassword
    }
}

function enableUser
{
    $ADSIUser.userFlags = $EnableUser
    $ADSIUser.setInfo()
    "Enabled $username"
}

function disableUser
{
    $ADSIUser.userFlags = $DisableUser
    $ADSIUser.setInfo()
    "Disabled $username"
}

$ADSIUser = [adsi]"WinNT://$env:COMPUTERNAME/$username" 

setPassword($password)

$EnableUser = 512
$DisableUser = 2

if($enable){enableUser}
elseif($disable){disableUser}
