#Requires -runasadministrator
#
# Deletes files from various locations used for temporary storage

$tempFolders = @( “C:\Windows\Temp\*”, “C:\Documents and Settings\*\Local Settings\temp\*”, “C:\Users\*\Appdata\Local\Temp\*”,"C:\windows\ccmcache\*")

Remove-Item $tempFolders -recurse -force -ErrorAction SilentlyContinue
