##########################################################################
#
# This script is designed to be used as a function, and not as a standalone
# script. It will find the index of the nth instance of a charatcer or
# string. 
# 
# Example: if you are looking for the 3rd e in the string "The bee's knees"
# it would be used as such: indexOfNth("The bee's knees", "e", 3)
##########################################################################

function indexOfNth([string]$string, [string]$find, [int]$N)
{
    #if index doesn't exist, initialize it to 0
    try
    {
        Get-Variable index -ErrorAction silentlyContinue | Out-Null
    }
    Catch
    {$index = 0}

    #If it's not the last run
    if($N -ne 1)
    {
        #pass the next index into the function
        $index = $string.indexOf($find,$index)+1
        indexofNth $string $find ($N-1)
    }
    #If last run
    else
    {
        #find last index
        $index = $string.indexOf($find,$index)
        #Return result
        $index
        Remove-Variable index
    }    
}
