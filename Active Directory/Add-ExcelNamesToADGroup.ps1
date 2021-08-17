<#
.SYNOPSIS
    Reads usernames from an Excel document and adds them to an AD group.
    
.DESCRIPTION
    Requires an excel document which has a list of usernames in column A to read from. 
    The script will read all usernames and add them to the given AD group.
#>

$excelPath = Read-Host -Prompt "Location of Excel Sheet"
$group = Read-Host -Prompt "AD Group name to add users to"

# Open Excel and get workbook
$excel = New-Object -Com Excel.Application
$ids = ((get-process excel | select MainWindowTitle, ID, StartTime | Sort StartTime)[-1]).Id
$wb = $excel.Workbooks.Open($excelPath)


# Get excel worksheet
$sh = $wb.Sheets.Item(1)
# Find highest numbered used row
[int]$endRow = $sh.UsedRange.rows.count
# Read each row and add username to group
for( $row = 1; $row -le $endRow; $row++) {
    $name = $sh.Cells.Item($row, 1).Value2
    Add-ADGroupMember -Identity $group -Members $name
    if($?) {
        "$name added to $group"
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