# If user used "Run with Powershell" command from right-click menu
$Command = (((get-itemproperty -literalpath 'HKLM:\SOFTWARE\Classes\SystemFileAssociations\.ps1\Shell\0\Command').'(default)').split('"').trim() | where {$_})[2] -replace " '%1'",""

if($MyInvocation.line.StartsWith($command))
{
    Start-Sleep 30
}