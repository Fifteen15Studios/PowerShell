<#
.SYNOPSIS
    Reads usernames from an Excel document and adds them to an SCCM Collection.
    
.DESCRIPTION
    Requires an excel document which has a list of usernames in column A to read from. 
    The script will read all usernames and add them to the given SCCM collection.
#>

$SiteCode = "" # Site code for your SCCM site
$domainName = ""

#-------------------------Start Script-----------------------------------------
$excelPath = Read-Host -Prompt "Location of Excel Sheet"
$collection = Read-Host -Prompt "Collection to add users to"

#Save current path
$path = (Get-Location).Path

# Connect to SCCM Drive
Import-Module "C:\Program Files (x86)\Microsoft Endpoint Manager\AdminConsole\bin\ConfigurationManager.psd1"
Set-Location ("$siteCode"+":")

# Open Excel and get workbook
$excel = New-Object -Com Excel.Application
$ids = ((get-process excel | select MainWindowTitle, ID, StartTime | Sort StartTime)[-1]).Id
$wb = $excel.Workbooks.Open($excelPath)

# Get excel worksheet
$sh = $wb.Sheets.Item(1)
# Find highest numbered used row
[int]$endRow = $sh.UsedRange.rows.count
# Read each row and add username to collection
for( $row = 1; $row -le $endRow; $row++) {
    $name = $sh.Cells.Item($row, 1).Value2
    Get-CMUser -name "$domainName\$name" | Add-CMUserCollectionDirectMembershipRule -CollectionName $collection
    if($?) {
        "$domainName\$name added to $collection"
    }
}
    
set-location $path

# Close Excel
$excel.WorkBooks.Close()
$excel.Quit()

#Force Excel instance to close
Stop-Process -id $ids[0]

#Cleanup
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()