param (
[parameter(Mandatory=$true, HelpMessage="Name of computer to transfer files to")]
[string ]$DestinationName
)

$SiteCode = "<site code>" # Must be hard coded. Detection methods don't seem to work
$USMTSourceCollection = "USMT Scan" # Change to the name of your collection
$USMTDestCollection = "USMT Load" # Change to the name of your collection

#Save current path
$path = (Get-Location).Path

#------Connect to SCCM Drive------
Import-Module $env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1

$ProviderMachineName = "ServerName.FQDN.COM" # Site Server

#Service Account
$Password = ConvertTo-SecureString "PlainTextPassword" -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential ("Username", $Password)

$output = new-psdrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $ProviderMachineName -Description "$SiteCode Primary Site" -Credential $creds
if($output -ne $null)
{"Connected to SCCM"}

Set-Location ("$siteCode"+":")
#-----Connected to SCCM Drive------

$source = Get-CMDevice -Name $env:COMPUTERNAME
$dest = Get-CMDevice -Name $DestinationName

#If both computers were found
if($source -ne $null -and $dest -ne $null)
{
    $output = New-CMComputerAssociation -SourceComputer $source.name -DestinationComputer $dest.Name -MigrationBehavior CaptureAndRestoreAllUserAccounts
    "Computer Association created"

    Add-CMDeviceCollectionDirectMembershipRule -CollectionName $USMTSourceCollection -Resource $source
    "$($source.name) added to $USMTSourceCollection"
    Add-CMDeviceCollectionDirectMembershipRule -CollectionName $USMTDestCollection -Resource $dest
    "$($dest.name) added to $USMTDestCollection"
}
else
{
    if($source -eq $null)
    {"Device $env:COMPUTERNAME not found"}
    else
    {"Device $DestinationName not found"}
}

#return to original path
Set-Location $path

#Run machine policy update, so that the Task Sequence shows in Software Center
try{
    $output = Invoke-WmiMethod -Namespace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"
    if($output -ne $null)
    {"Triggered machine policy update"}
}
catch{
    
}

#If destination computer is on, force machine policy there too
if (Test-Connection -ComputerName $($dest.name) -Quiet)
{
    try{
        Invoke-Command -ComputerName $($dest.name) -Credential $creds -ScriptBlock {
        $output = Invoke-WmiMethod -Namespace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"
        if($output -ne $null)
        {"Triggered machine policy update on destination PC"}
        }
    }
    catch
    {
    } 
}