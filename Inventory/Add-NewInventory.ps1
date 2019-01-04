# This script reads from an excel spreadsheet - $InputFile - looking for PC 
# computer names then looks for that computer name within 2 tabs of 
# $InventoryFile. Those 2 tabs - $InventoryWorkSheet and $SurplusWorkSheet - 
# contain current inventory and items sent to surplus, respecitvely.
#
# If the computer name read from $InputFile is not found in either tab of 
# $InventoryFile, the entire row (line) from $InputFile is copied into the
# Inventory tab of $InventoryFile

param( 
    [ValidateScript({
            if($_ -notmatch "(\.xls|\.xlsx)"){
                throw "The file specified in the InputFile argument must be of type xls or xlsx."
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

$NameColumn = $InputWorkSheet.Cells.Find("Computer Name").column
$OSColumn = $InputWorkSheet.Cells.Find("Operating System").column

#Find number of rows
$endRow = $InputWorkSheet.UsedRange.SpecialCells(11).Row
$endInventoryRow = $InventoryWorkSheet.UsedRange.SpecialCells(11).Row + 1
$endSurplusRow = $SurplusWorkSheet.UsedRange.SpecialCells(11).Row + 1

#Read Rows
for($i = 1; $i -le $endRow; $i++)
{
    $row = $InputWorkSheet.Rows[$i]
    $Name = $row.value2[1,$NameColumn]
    $OS = $row.Value2[1,$OSColumn]
    $Found = $null

    #if name and OS are both not empty, and not header row
    if(($Name -ne "" -and $Name -ne $null -and $name -ne "Computer Name") -and
        ($OS -ne "" -and $OS -ne $null))
    {
        $Found = $InventoryWorkSheet.Cells.Find($Name)

        #If not in Inventory, check surplus
        if(-not $found)
        {
            $Found = $SurplusWorkSheet.Cells.Find($Name)

            #If not in Surplus either
            if(-not $Found)
            {
                #Add to inventory
                $InputWorkSheet.Rows[$i].copy() | Out-Null
                $Range = $InventoryWorkSheet.Range("A" + $endInventoryRow)
                $InventoryWorkSheet.paste($Range)
                $endInventoryRow++
            }
        }
    }
    
    Write-Progress -Activity "Checking for new inventory..." -PercentComplete (($i/$endRow)*100)

    #Every 100 items, display status
    <#if($i % 100 -eq 0)
    {
        "$i/$endRow"
    }#>
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
#$InventoryWorkBook.SaveAs($InventoryFile)

$InputWorkBook.Close()
$InventoryWorkBook.Close()

$excel.Workbooks.Close()
$excel.Quit()

#Cleanup
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

#Force Excel app to close
while( [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)) {}
