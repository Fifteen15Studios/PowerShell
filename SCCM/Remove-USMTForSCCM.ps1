# Removes USMT computer association, and removes the computers from the SCCM
# collection that the USMT task sequences are deployed to.
#
# This script requires an SCCM service account to connect to the SCCM 
# environment in the PowerShell window.

$SiteCode = "<site code>" # Must be hard coded. Detection methods don't seem to work
$USMTSourceCollection = "USMT Scan"  # Change to the name of your collection
$USMTDestCollection = "USMT Load"  # Change to the name of your collection

#Save current path
$path = (Get-Location).Path

#------Connect to SCCM Drive------
Import-Module $env:SMS_ADMIN_UI_PATH\..\ConfigurationManager.psd1

$ProviderMachineName = "ServerName.FQDN.COM" # Site Server

#Service Account
$Password = ConvertTo-SecureString "Password" -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential ("Username", $Password)

$output = new-psdrive -Name $SiteCode -PSProvider "AdminUI.PS.Provider\CMSite" -Root $ProviderMachineName -Description "$SiteCode Primary Site" -Credential $creds
if($output -ne $null)
{"Connected to SCCM"}

Set-Location ("$siteCode"+":")
#-----Connected to SCCM Drive------

$DestinationName = $env:COMPUTERNAME
$USMTAssociation = Get-CMComputerAssociation -DestinationComputer $DestinationName

#Remove computers from collections
Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $USMTSourceCollection -ResourceName $($USMTAssociation.SourceName) -force
"$($USMTAssociation.SourceName) removed from $USMTSourceCollection"
Remove-CMDeviceCollectionDirectMembershipRule -CollectionName $USMTDestCollection -ResourceName $DestinationName -force
"$DestinationName removed from $USMTDestCollection"

#Remove USMT Association
$USMTAssociation | Remove-CMComputerAssociation -Force
"Removed computer association"

#Run machine policy update, so that the Task Sequence shows in Software Center
try{
    $output = Invoke-WmiMethod -Namespace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"
    if($output -ne $null)
    {"Triggered machine policy update"}
}
catch
{
    
}

#If source computer is on, force machine policy there too
if (Test-Connection -ComputerName $($USMTAssociation.SourceName) -Quiet)
{
    try{
        Invoke-Command -ComputerName $($dest.name) -Credential $creds -ScriptBlock {
        $output = Invoke-WmiMethod -Namespace "Root\CCM" -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000021}"
        if($output -ne $null)
        {"Triggered machine policy update on source PC"}
        }
    }
    catch
    {

    }
}
