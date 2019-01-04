# This script reads from an excel spreadsheet - $InputFile - looking for PC
# serial numbers. For each serial number, it looks for that serial number 
# within the inventory tab - $inventoryWorkSheet - of $InventoryFile.
#
# If a serial number in $InputFile is found within the inventory tab, the 
# entire row (line) is copied from the inventory tab, pasted in the surplus
# tab, then removed from the inventory tab.

param( 
    [ValidateScript({
            if($_ -notmatch "(\.xls|\.xlsx)"){
                throw "The file specified in the OutputFile argument must be of type xls or xlsx."
            }
            return $true 
        })]
    [parameter(ValueFromPipeline=$true, Mandatory=$true, position=0, HelpMessage="Location of the input file")]
    [string]$InputFile
)

#----------------------------------Start Setup---------------------------------
$InventoryFile = "" #Make sure to put in the name of an Excel file (.xls or .xlsx)

#If no full path provided
if((-not $InputFile.StartsWith("\\")) -and (-not $InputFile.Substring(1,2).Equals(":\")))
{
    $InputFile = "$((Get-Location).path)\$InputFile"
}


#Create Excel Object
try{
    $excel = New-Object -ComObject Excel.Application
}
catch [System.Runtime.InteropServices.COMException]
{
    Write-Error -Message "Can't create Excel object. Is Excel installed?" -Category NotInstalled -RecommendedAction "Check to ensure that Excel is installed."
    exit
}
catch
{
    Write-Error "An error occurred while trying to create an Excel object: `n$($_.Exception.Message)"
    exit
}


$InputWorkBook = $excel.Workbooks.Open($InputFile)
$InputWorkSheet = $InputWorkBook.Sheets.Item(1)

$InventoryWorkBook = $excel.Workbooks.Open($InventoryFile)
$InventoryWorkSheet = $InventoryWorkBook.Sheets.Item(1)
$SurplusWorkSheet = $InventoryWorkBook.Sheets.Item(2)

#------------------------------------End Setup---------------------------------

#-------------------------------------Do Work----------------------------------

#Find number of rows
$endRow = $InputWorkSheet.UsedRange.SpecialCells(11).Row
$FindColumn = $InputWorkSheet.Cells.Find("Serial").column

$endSurplusRow = $SurplusWorkSheet.UsedRange.SpecialCells(11).Row + 1

#Read Rows
for($i = 1; $i -le $endRow; $i++)
{
    $row = $InputWorkSheet.Rows[$i]
    $serial = $row.value2[1,$FindColumn]

    if($serial -ne "" -and $serial -ne $null)
    {
        $Found = $InventoryWorkSheet.Cells.Find($Serial)
        if($found)
        {
            #Copy to other tab
            $InventoryWorkSheet.Rows[$Found.Row].copy() | Out-Null
            $Range = $SurplusWorkSheet.Range("A" + $endSurplusRow)
            $SurplusWorkSheet.paste($range)
            $endSurplusRow++
            #Delete
            $InventoryWorkSheet.Rows[$Found.Row].Delete() | Out-Null
        }
    }

    Write-Progress -Activity "Processing Inventory..." -PercentComplete (($i/$endRow)*100)

}



#-----------------------------------Start Cleanup------------------------------

#adjust the column width so all data's properly visible 
$InventoryWorkSheet.UsedRange.EntireColumn.AutoFit() | Out-Null

#if already exists, delete so it can be overwritten
#if(Test-Path $InventoryFile)
#{Remove-Item $InventoryFile}

#Rename file with today's date
$InventoryFile = $InventoryFile.Substring(0,$InventoryFile.LastIndexOf("."))
$InventoryFile += "-$(Get-Date -Format "yyyy-MM-dd").xlsx"

#Save the file
$excel.ActiveWorkbook.SaveAs($InventoryFile)

#Close the files and exit Excel
$InputWorkBook.Close()
$InventoryWorkBook.Close()

$excel.Workbooks.Close()
$excel.Quit()

#Cleanup
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

#Force Excel app to close
while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)) {}
