# Outputs all task sequence variables and their value to $location
#
# This script should be called using the "Run PowerShell Script" task sequence
# step. That step will need to point to a package with a source set to a folder
# containing this .ps1 file.
#
# The script will fail if run from outisde of a Task Sequence, as the 
# 'New-Object' line will not find what it needs to create the object.

#Where to output the variables to
$Location = "$ENV:TEMP\TS_Vars.txt"

#Get Task Sequence object
$tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment

$Output = ""

#Read each vaiable and put it in a string variable
Foreach( $variable in $tsenv.GetVariables() )
{ $Output += $variable + " = " + $tsenv.value($variable) + "`n" }

#Write to file
$Output >> $Location
