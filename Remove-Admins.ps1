#Requires -runasadmin
#
# Script:   Remove-Admins.ps1
#
# Purpose:  Removes specified user(s) from local administrator group 
# of a Windows PC
#
# Parameters: -UserName - Required - Username(s) to remove from the admin group

param(
    [Parameter(ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true, mandatory = $true)]
    [string[]]$UserName
    )

#Get group object
$obj_group = [ADSI]"WinNT://localhost/Administrators,group"

#Get members of group
$members= @($obj_group.psbase.Invoke("Members")) | foreach{(([ADSI]$_).InvokeGet("Name")).tolower()}

$members | foreach{
 
    #Check to see if these users are in the group. If so, remove them.
    if($UserName -contains $_)
    {
        try
        {
            $obj_group.Remove("WinNT://$_")
            Write-Host "Removed ""$_"" from ""Administrators"" group" -ForegroundColor Green
        }
        catch
        {
            Write-Host "Failed to remove ""$_"" from ""Administrators"" group" -ForegroundColor Red
        }
    }
}
