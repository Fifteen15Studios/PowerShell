###############################################################################
#
# This script will generate a random string with the options of that string
#   represented by the parameters that are used when calling the script. By
#   default the string contains only capital and lowercase letters.
#
# Parameters:
#   Length: int - how many characters long the string should be
#   Numbers: Switch - If used, numbers may be included in the string
#   Symbols: Switch - If used, some symbols may be included in the string
#
# Note: There has been no work done to prevent similar characters from 
#   appearing in the string. For example, 1, l, and I look very similar in some
#   fonts. A string may be generated which contains l and I in the same string.
#
###############################################################################

param(
    [Parameter(mandatory=$true)]
    [int]$Length,
    [Parameter(mandatory=$false)]
    [switch]$Numbers,
    [Parameter(mandatory=$false)]
    [switch]$Symbols
)

# First range is capital letters, second range is lower case letters
# All numbers are ASCII values
$Range = (65..90) + (97..122)

if($Numbers)
{
    $Range += (48..57)
}
if($Symbols)
{
    $Range += (33..47)+(58..64)
}

# Create and display password
-join ($Range | Get-Random -Count $Length | % {[char]$_})
