###############################################################################
# 
# This function finds a column letter in Excel based in the column number.
#   For example, column 1 is also refered to as column A. Column 26 is column Z
#   so column 27 is AA. Column 702 is column ZZ, which makes column 703 AAA
# 
# Input: Column number. Ex: 27
# Output: Column Letter. Ex: AA
# 
###############################################################################
function columnNumToLetter([int]$col) {

    $output = ""
    
    # If value is 2 or more letters
    if($col -gt 26) {
        #Find out if value is more than 2 letters
        $multiplier = [math]::floor(($col -1) / 26)
        
        # Continue while value is more than 2 letters
        while($multiplier -gt 26) {
            # Output current letter, starting with most significant
            $output += columnNumToLetter $multiplier
            # find next letter value
            $multiplier -= (26 * $multiplier)
        }
        # At this point, there are 2 letters that remain to be displayed
        # multiplier holds the most first letter. Output it
        $output += columnNumToLetter $multiplier
        
        # Find second letter
        if($col %26 -ne 0)
            {$col = $col % 26}
        else
            {$col = 26}

    }

    # Add Letter to output
    if($col -gt 0 -and $col -le 26) {
        $output += switch ($col)
        {
            1 {"A"}
            2 {"B"}
            3 {"C"}
            4 {"D"}
            5 {"E"}
            6 {"F"}
            7 {"G"}
            8 {"H"}
            9 {"I"}
            10 {"J"}
            11 {"K"}
            12 {"L"}
            13 {"M"}
            14 {"N"}
            15 {"O"}
            16 {"P"}
            17 {"Q"}
            18 {"R"}
            19 {"S"}
            20 {"T"}
            21 {"U"}
            22 {"V"}
            23 {"W"}
            24 {"X"}
            25 {"Y"}
            26 {"Z"}
        }
    }

    # Return the final value
    $output
}
