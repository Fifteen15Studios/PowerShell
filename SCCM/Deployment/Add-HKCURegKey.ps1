<#
.SYNOPSIS
    Adds a key to the HKCU section of the registry
 
 .DESCRIPTION
    This script will find the logged in user and add a key to their section 
    of the HKU portion of the registry.
    If you would like to add a key for all users who log on to the computer,
    use the "allUsers" switch, and provide a unique identifier in the
    "AllUsersUniqueKey" parameter.

    Acceptable values for -type include:
    "string" or "REG_SZ"
    "expandstring" or "REG_EXPAND_SZ"
    "binary" or "REG_BINARY"
    "dWord" or "REG_DWORD"
    "multistring" or "REG_MULTI_SZ"
    "qword" or "REG_QWORD"
    Default value is "DWORD"
 
.LINK
    
.NOTES
    Version: 1.0
 
.EXAMPLE
    Add-HKCURegKey "\Software\Example" "Activated" "True" "string"
.EXAMPLE
    Add-HKCURegKey "\Software\MySoftware" "MyKey" "100" "dword" -allusers "MySoftware"
#>
[CmdletBinding(DefaultParameterSetName = 'Basic')]
param(
    [parameter(ParameterSetName = "AllUsers", Mandatory = $true)]
    [parameter(ParameterSetName = "Basic", Mandatory = $true)]
    [String]$EndPath,
    [parameter(ParameterSetName = "AllUsers", Mandatory = $true)]
    [parameter(ParameterSetName = "Basic", Mandatory = $true)]
    [String]$KeyName,
    [parameter(ParameterSetName = "AllUsers", Mandatory = $true)]
    [parameter(ParameterSetName = "Basic", Mandatory = $true)]
    $KeyValue,
    [parameter(ParameterSetName = "AllUsers")]
    [ValidateSet("string", "REG_SZ", "expandstring", "REG_EXPAND_SZ", "binary", "REG_BINARY", "dWord", "REG_DWORD", "multistring", "REG_MULTI_SZ", "qword", "REG_QWORD")]
    [parameter(ParameterSetName = "Basic")]
    [ValidateSet("string", "REG_SZ", "expandstring", "REG_EXPAND_SZ", "binary", "REG_BINARY", "dWord", "REG_DWORD", "multistring", "REG_MULTI_SZ", "qword", "REG_QWORD")]
    [String]$Type = "DWORD",
    [parameter(ParameterSetName = "AllUsers")]
    [switch]$AllUsers,
    [parameter(ParameterSetName = "AllUsers", mandatory = $True)]
    [String]$AllUsersUniqueKey
)

# Use "$((getUserRegPath).regpath)" to get reg path in a string
# "$((getUserRegPath).user)" to get username in a string
function getUserRegPath() {
        
    # Get all logged in users, even if they're logged in using RDP
    $Users = quser.exe /server:$computer 2>$null | select -Skip 1

    # Filter the results
    $loggedOnUsers = foreach ($user in $users){
        ((($user).Split(" ",2)[0]).split(">",2)[1]).trim()
    }

    #map psdrive for hku
    if(-not (Test-Path "HKU:")) {
        New-PSDrive -Name hku -PSProvider Registry -Root hkey_users -Scope Global | Out-Null
    }

    #find matching volatile env
    $environment =  foreach ($user in $loggedOnUsers) {
        [PSCustomObject]@{
            user = $user
            regpath = Join-Path 'hku:' (((Get-ItemProperty 'hku:\*\Volatile Environment' | Where-Object username -eq $user).pspath -split '\\')[2])
        }
    }

    $environment
}

switch($Type.ToLower()) {
    string {$addType="REG_SZ"}
    expandstring {$addType="REG_EXPAND_SZ"}
    binary {$addType="REG_BINARY"}
    dWord {$addType="REG_DWORD"}
    multistring {$addType="REG_MULTI_SZ"}
    qword {$addType="REG_QWORD"}
}

if($AllUsers) {

    if(-not (Test-Path "HKLM:\Software\Microsoft\Active Setup\Installed Components\$AllUsersUniqueKey")) {
        New-Item "HKLM:\Software\Microsoft\Active Setup\Installed Components\$AllUsersUniqueKey" -Force
    }

    # Adds key to all user's HKCU area when they log in
    New-ItemProperty -Path "HKLM\Software\Microsoft\Active Setup\Installed Components\$AllUsersUniqueKey" -Name StubPath -Value "reg add '\`"HKCU\$EndPath`\`" /v $KeyName /d $keyValue /t $addType /f" -Force | Out-Null
}

# Adds key to current users' HKCU area
$Paths = (getUserRegPath).regpath
foreach ($path in $paths) {
    # If path doesn't exist, create it
    if(-not (Test-Path "$path$endpath")) {
        New-Item -Path "$path$endpath" -Force
    }

    # Create key
    New-ItemProperty -Path "$Path$endpath" -Name $KeyName -Value $keyValue -PropertyType $Type -Force | Out-Null
}

Remove-PSDrive hku | Out-Null