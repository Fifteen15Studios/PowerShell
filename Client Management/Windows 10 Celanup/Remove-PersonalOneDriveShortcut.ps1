###############################################################################
#
# This script removes the "OneDrive" shortcut on the left side of File Explorer
#
# Since many people do not use this shortcut, it serves only as a source of
# confustion. Especially if you instead use an enterprise OneDrive.
#
###############################################################################

$ID = "{018D5C66-4533-4307-9B53-224DE2ED1FE6}"

#If HKCR no accessible
if( -not (test-path HKCR:))
{
    #Map HKCR as PS drive
    New-PSDrive -Name HKCR -PSProvider Registry -Root HKEY_CLASSES_ROOT
}

Set-ItemProperty -Path "HKCR:\CLSID\$ID" -Name "System.IsPinnedToNameSpaceTree" -Value 0 -Force
