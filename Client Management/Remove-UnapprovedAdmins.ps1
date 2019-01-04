#Purpose:  Removes unauthorized user(s) from local administrator group
#
#Requires -runasadmin

# Add other values here, each entryseparated by a coma
$Approved = "Administrator","Domain Admins" 

#Get group object
$obj_group = [ADSI]"WinNT://localhost/Administrators,group"

#Get members of group
$Members= @($obj_group.psbase.Invoke("Members")) | foreach{(([ADSI]$_).InvokeGet("Name"))}

$Members | foreach{
 
    #Check to see if any unacceptable users are in the group. If so, remove them.
    if(-not ($Approved -contains $_))
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
