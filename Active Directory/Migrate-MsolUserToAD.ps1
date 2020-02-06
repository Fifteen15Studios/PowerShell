###############################################################################
#
# This script migrates users from Office 365 to local active directory. All
#    necessary attributes of the user will be coppied from Office 365 to ensure
#    that the user properly syncs between systems. In order to ensure this 
#    syncronization, the following attriubutes must be the same in both systems:
#    Display Name, Principal Name, Mail, ProxyAddresses
#
# In order for the script to actually execute the migration, it must be called
#    with the -ForReal switch. If this switch is not present, it will simply
#    display what user would be migrated in a real scenario.
#
# Parameters:
#    ForReal - Perform the migration. Without this switch no changes are made
#    RandomPassword - Create a randomized password for each migrated user. This
#        password will be emailed to the user, and displayed on screen after 
#        the script finishes executing.
#
###############################################################################

#Requires -modules MSOnline, ActiveDirectory

[CmdletBinding(DefaultParametersetName='None')]
param(
    [Switch]$ForReal,
    [Parameter(ParameterSetName="Random")]
    [Switch]$RandomPassword,
    # Used as "From" address when sending randomized password to users
    # Make sure this is a real address in your domain or your email might get filtered into spam
    [Parameter(ParameterSetName="Random", Mandatory=$true)]
    [ValidateScript({
        if($_ -notmatch ".*@.*\..*")
        {throw "Invalid Email Address"}
        else
        {return $true}
    })]
    $EmailFrom
)

Import-Module ActiveDirectory

#------------------------------------------------------------------------------
# Feel free to modify the variables in this area if necessary

# Get values from domain controller
$OUPath = (Get-ADDomain).UsersContainer
$DomainFull = (Get-ADDomain).DNSRoot

$DomainPrefix = $DomainFull.Substring(0, $DomainFull.IndexOf("."))
$DomainSuffix = $DomainFull.Substring($DomainFull.IndexOf(".")+1)

# Change these if your email suffix is different from your domain name
# For example, if your domain is a contoso.local domain, but your email address ends with @contoso.com
# In this case, EmailSuffix should be contoso.com and SmtpAddress should start with contoso-com
# Or you can use your own SMTP server
$EmailSuffix = $DomainFull
$SmtpAddress = "$DomainPrefix-$DomainSuffix.mail.protection.outlook.com"

# Don't modify below this line
#------------------------------------------------------------------------------

$ErrorActionPreference = "Stop"

# Generate a random password of given length
function randomPassword()
{
    param(
        [Parameter(mandatory=$true)]
        [int]$length,
        [Parameter(mandatory=$false)]
        [switch]$Numbers,
        [Parameter(mandatory=$false)]
        [switch]$Symbols
    )

    # Letters only
    $range = (65..90) + (97..122)

    if($Numbers)
    {
        $range += (48..57)
    }
    if($Symbols)
    {
        $range += (33..47)+(58..64)
    }

    # Letters, numbers, and symbols
    return -join ($range| Get-Random -Count $length | % {[char]$_})
}

Import-Module MSOnline

# If not connected to MS Online yet
if(-not (Get-MsolDomain -ErrorAction SilentlyContinue))
{
    $O365Cred = Get-Credential -Message "Please provide credentials for Microsoft Online"

    "Connecting to Microsoft Online..."
    
    $O365Session = New-PSSession –ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365Cred -Authentication Basic -AllowRedirection -ErrorAction Stop
    Connect-MsolService –Credential $O365Cred
    $ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Cred -Authentication Basic -AllowRedirection
    
    Import-PSSession $O365Session | Out-Null
}

# Get name from user
$InputName = Read-Host "Enter the name of the user you would like to migrate"

$Users = Get-MsolUser -SearchString $InputName

function Get-Password()
{
    # Genrate and display random password
    if($RandomPassword)
    {
        $Password = randomPassword 10 -Numbers -Symbols
        "Password: $Password"
    }
    # Get password from user
    else 
    {
        $Password = Read-Host "Enter a password for $($User.DisplayName) ($UserName)" -AsSecureString
    }

    return $Password
}

# Might have found multiple users
foreach($User in $Users)
{
    # Ask for each user if you would like to migrate
    $Answer = Read-Host "Would you like to migrate $($User.DisplayName) ($($User.UserPrincipalName))? (Y/[N])"

    if($Answer -ne "" -and $Answer.ToLower()[0] -eq "y")
    {
        "Getting User Info..."
        $FirstName = $User.FirstName #GivenName
        $LastName = $User.LastName #Surname
        $DisplayName = $User.DisplayName #Name
        $ProxyAddresses = $User.ProxyAddresses 
        $PrincipalName = $user.UserPrincipalName #UserPrincipalName

        # Make sure sign-in name is really email address
        if($User.SignInName -like "*@$EmailSuffix")
        {
            $EmailAddress = $User.SignInName
            $UserName = $EmailAddress.Substring(0, $EmailAddress.IndexOf("@")) #SamAccountName

            $Mail = (Get-User $UserName).WindowsEmailAddress

            # Only actually migrate the user if switch is set - as a safety net
            if(-not $ForReal)
            {
                # Check if user already exists
                try{
                    $RandomPassword = $true
                    Get-ADUser $UserName | Out-Null
                    Write-Host "$UserName already exists in Active Directory" -ForegroundColor Yellow
                }
                catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]{
                    $password = Get-Password

                    # If username doesn't exist, it will end up here. Which is good.
                    $NewADUser = New-Object -TypeName PSObject

                    $NewADUser | Add-Member -MemberType NoteProperty -Name FirstName $FirstName
                    $NewADUser | Add-Member -MemberType NoteProperty -Name LastName $LastName
                    $NewADUser | Add-Member -MemberType NoteProperty -Name DisplayName $DisplayName
                    $NewADUser | Add-Member -MemberType NoteProperty -Name PrincipalName $PrincipalName
                    $NewADUser | Add-Member -MemberType NoteProperty -Name UserName $UserName
                    $NewADUser | Add-Member -MemberType NoteProperty -Name Mail $Mail

                    $NewADUser
                }

                Write-Host "NOTE: No changes have been made. Run this script with -ForReal to perform the migration." -ForegroundColor Yellow
            }
            # Do the actual migration
            else
            {
                # Check if user already exists
                try{
                    Get-ADUser $UserName | Out-Null
                    Write-Host "ERROR: $Username already exists in Active Directory`r`nNo changes have been made." -ForegroundColor Red
                }
                # If user doesn't already exist, create it
                catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
                {
                    $password = Get-Password

                    if($Password -isnot [SecureString])
                    {
                        $PlainPassword = $Password
                        $Password = ConvertTo-SecureString -String $PlainPassword
                    }

                    "Creating $($User.DisplayName) ($UserName)..."

                    New-ADUser -SamAccountName $UserName -Name $DisplayName -GivenName $FirstName -Surname $LastName -UserPrincipalName $PrincipalName -Enabled $True -Path $OUPath -AccountPassword $Password
                    
                    # Get newly created user
                    $NewUser = (Get-ADUser -Filter "SamAccountName -eq ""$UserName""")

                    # Add Proxy Addresses
                    foreach($Address in $ProxyAddresses)
                    {
                        $NewUser.ProxyAddresses.add($Address)
                    }

                    # Send the randomized password to the user in an email
                    if($RandomPassword)
                    {
                        $MailAnswer = Read-Host "Email user info to $($Mail)? (Y/[N])"

                        if($MailAnswer -ne "" -and $MailAnswer.ToLower()[0] -eq "y")
                        {
                            Send-MailMessage -To $Mail -From $EmailFrom -SmtpServer $SmtpAddress -Subject "New Account Info" -Body "A new Active Directory account has been created for you. Below is the sign-in info for your new account.`r`n`r`nUsername: $UserName`r`nPassword: $PlainPassword"
                            "User info has been sent to $Mail"
                        }
                    }
                }
            }
        }
    }
}

$ErrorActionPreference = "Continue"
