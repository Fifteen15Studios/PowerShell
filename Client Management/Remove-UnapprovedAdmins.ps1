##############################################################################
#
# Purpose:  Removes unauthorized user(s) from local administrator group
#
###############################################################################

#Requires -runasadmin

$Approved = [System.Collections.ArrayList]@()
# Add other values here, each entry separated by a comma. Yes, the double parentheses are necessary.
$Approved.AddRange(("Administrator","Domain Admins"))

# Optional - Add all users of an AD group to the approved list.
#$group = (Get-ADGroupMember "Approved_Admins").name
#$Approved.AddRange($group)

# Get administrators group object
$obj_group = [ADSI]"WinNT://localhost/Administrators,group"

# Get members of group
$Members= @($obj_group.psbase.Invoke("Members")) | foreach{(([ADSI]$_).InvokeGet("Name"))}
$Count = 0

# Iterate through all users in the administrators group
$Members | foreach{
 
    # Check to see if any unacceptable users are in the group. If so, remove them.
    if(-not ($Approved -contains $_))
    {
        $count++

        try
        {
            $obj_group.Remove("WinNT://$_")
            Write-Host "Removed ""$_"" from ""Administrators"" group on $env:COMPUTERNAME" -ForegroundColor Green
        }        
        catch
        {
            Write-Host "Failed to remove ""$_"" from ""Administrators"" group on $env:COMPUTERNAME" -ForegroundColor Red
        }
    }
}

# Output if nothing was found
if($Count -eq 0)
{
    Write-Host "No Unathorized administrators found on $env:COMPUTERNAME!" -ForegroundColor Green
}
