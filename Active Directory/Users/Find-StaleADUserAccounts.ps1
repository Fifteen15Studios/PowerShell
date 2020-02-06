###############################################################################
#
# This script gets all users from Active Directory and find the 
#    LastLogonTimestamp for each user. This attribute is not synched between
#    Domain controllers on every synch, so the information provided may not be
#    exactly correct, but it will give a fairly good estimate of the last time
#    each user has logged onto a PC.
#
# In order to get precise information, the LastLogon attribute would need to be
#    pulled from each domain controller, and this process can be time consuming,
#    especially in larger environments.
#
# Parameters:
#    StaleDays - The script will look for accounts that have not been logged 
#        onto in this number of days (or more) and mark any such accounts as 
#        stale
#    CreateCSV - If set, data will be put into a CSV file
#    CsvLocation - Where to put the CSV file
#
###############################################################################

[CmdletBinding(DefaultParametersetName='None')]
param(
    [Int]$StaleDays = 90,
    [Parameter(ParameterSetName="CSV")]
    [Switch]$CreateCSV,
    [Parameter(ParameterSetName="CSV", Mandatory = $true)]
    [ValidateScript({
            if( -Not ($_ | Split-Path | Test-Path) ){
                throw "folder does not exist"
            }
            if($_ -notmatch "(\.csv)"){
                throw "The file specified in the InputFile argument must be of type csv."
            }
            return $true
        })]
    [System.IO.FileInfo]$CSVLocation

)

Import-Module ActiveDirectory

$users = get-aduser -f * -pr lastLogonTimestamp | select Name, SamAccountName, lastLogonTimestamp | sort SamAccountName

foreach ($user in $users){
    # If timestamp exists
    if($user.lastLogonTimestamp -ne $null)
    {
        #Convert timestamp from integer into date
        $user.lastLogonTimestamp = [DateTime]::FromFiletime([Int64]::Parse($user.lastLogonTimestamp))

        $days = ((Get-Date) - $user.lastLogonTimestamp).days
        $user | Add-Member -MemberType NoteProperty -Name "Days since last login" -Value $days
        
        if($days -ge $StaleDays)
        {$user | Add-Member -MemberType NoteProperty -Name "Stale" -Value $true}
        else
        {$user | Add-Member -MemberType NoteProperty -Name "Stale" -Value $false}
    }
    # If timestamp doesn't exist
    else
    {
        $user.lastLogonTimestamp = "Never"
        $user | Add-Member -MemberType NoteProperty -Name "Days since last login" -Value ""
        $user | Add-Member -MemberType NoteProperty -Name "Stale" -Value "Unknown"
    }
}

if($CreateCSV)
{$users | Export-Csv -Path $CSVLocation -NoTypeInformation -Force}
else
{$users}
