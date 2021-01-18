Function Get-Something {
<#
.SYNOPSIS
    This is a basic overview of what the script is used for..
 
 .DESCRIPTION
    Describes in more detail what you're doing
 
.LINK
    https://thesysadminchannel.com/powershell-template -
    
# The 2 below won't actually show in Get-Help

.NOTES
    Name: Get-Something
    Author: theSysadminChannel
    Version: 1.0
    DateCreated: 2020-Dec-10
 
.EXAMPLE
    Get-Something -UserPrincipalName "username@thesysadminchannel.com"
#>
 
    [CmdletBinding()]
    param(
        [Parameter(
            Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0
            )]
        [string[]]  $UserPrincipalName
    )
 
    BEGIN {}
 
    PROCESS {}
 
    END {}
}
