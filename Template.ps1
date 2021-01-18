Function Get-Something {
<#
.SYNOPSIS
    This is a basic overview of what the script is used for..
 
 
.NOTES
    Name: Get-Something
    Author: theSysadminChannel
    Version: 1.0
    DateCreated: 2020-Dec-10
 
 
.EXAMPLE
    Get-Something -UserPrincipalName "username@thesysadminchannel.com"
 
 
.LINK
    https://thesysadminchannel.com/powershell-template -
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
