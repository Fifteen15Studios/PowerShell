#Requires -runasadministrator
#
# Deletes files from various locations used for temporary storage

$tempFolders = @( 
“C:\Windows\Temp\*”, # Default Windows temp location
“C:\Documents and Settings\*\Local Settings\temp\*”, # Windows XP user temp
“C:\Users\*\Appdata\Local\Temp\*”, # Windows Vista+ user temp
"C:\windows\ccmcache\*" # Default SCCM temp location
)

Remove-Item $tempFolders -recurse -force -ErrorAction SilentlyContinue
