<#
.Synopsis
   Function to copy group memberships from a source user to target users.
.DESCRIPTION
   Function to copy group memberships from a source user to multiple target users. It's
   also possible to make an exact duplicate by removing the existing memberships
   from the target accounts.
   Requires ActiveDirectory module.
.EXAMPLE
   Copy-GroupMembership -Source s.user -Targets t.user1,t.user2
   Adds the group memberships of user s.user to target users t.user1 and t.user2.
.EXAMPLE
   Copy-GroupMembership -Source s.user -Targets t.user1,t.user2 -RemoveExisting
   Adds the group memberships of user s.user to target users t.user1 and t.user2 and
   removes the existing group memberships from those target accounts, resulting in an
   exact duplicate.
.NOTES
   Author: Michaja van der Zouwen
.LINK
   https://itmicah.wordpress.com/2014/11/27/copy-ad-group-memberships-from-a-source-user-to-other-users

#>
function Copy-GroupMembership
{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param
    (
        # Source user account name
        [Parameter(Mandatory=$true,
                    Position=0)]
        [string]
        $Source,

        # Comma seperated list of target user accounts
        [Parameter(Mandatory=$true,
                    Position=1)]
        [string[]]
        $Targets,

        # Remove existing group memberships from target accounts
        [switch]
        $RemoveExisting
    )
    
    Write-Verbose 'Retrieving source group memberships.'
    $SourceUser = Get-ADUser $Source -Properties memberOf -ea 1
    
    foreach ($Target in $Targets) {

        Write-Verbose "Get group memberships for '$Target'."
        $TargetUser = Get-ADUser $Target -Properties memberOf

        If (!$TargetUser) {
            Write-Warning "Unable to find useraccount '$Target'. Skipping!"
        }
        else {
            # Hash table of source user groups.
            $List = @{}

            #Enumerate direct group memberships of source user.
            ForEach ($SourceDN In $SourceUser.memberOf)
            {
                # Add this group to hash table.
                $List.Add($SourceDN, $True)
                # Bind to group object.
                $SourceGroup = [ADSI]"LDAP://$SourceDN"

                Write-Verbose "Checking if '$target' is already a member of '$sourceDN'."
                If ($SourceGroup.IsMember("LDAP://" + $TargetUser.distinguishedName) -eq $False)
                {
                    if ($pscmdlet.ShouldProcess($Target, "Add to group '$SourceDN'"))
                    {
                        Write-Verbose "Adding '$target' to this group."
                        Add-ADGroupMember -Identity $SourceDN -Members $Target
                    }
                }
                else
                {
                    Write-Verbose "'$Target' is already a member of this group."
                }
            }

            #If required remove existing memberships
            If ($RemoveExisting) 
            {
                Write-Verbose 'Entering removal phase.'

                # Enumerate direct group memberships of target user.
                ForEach ($TargetDN In $TargetUser.memberOf)
                {
                    Write-Verbose "Checking if '$Target' is a member of '$TargetDN'."
                    If ($List.ContainsKey($TargetDN) -eq $False)
                    {
                        if ($pscmdlet.ShouldProcess($Target, "Remove from group '$TargetDN'"))
                        {
                            # Source user not a member of this group.
                            Write-Verbose "Removing '$Target' from this group."
                            Remove-ADGroupMember $TargetDN $Target
                        }
                    }
                    else
                    {
                        Write-Verbose "'$Target' is not a member of this group."
                    }
                } # end foreach
            } # end If
        } # end If-else
    }  # end foreach Target
} # end function
