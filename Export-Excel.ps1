Function Export-Excel{
<#
    .SYNOPSIS
        Accepts input of a PSObject collection of items and an Excel Object to create a worksheet 
        and populate it with the contents of the collection.
    .DESCRIPTION
        Input Info
    .NOTES
        Filename: 
        Author: Rick Wilcox 
        Requirements: 
    .LINK
        Input info  
    .PARAMETER  options
        Input Info
    .PARAMETER  displayProperty
        Input Info
    .PARAMETER  title
        Input Info
    .PARAMETER  mode
        Input Info
    .PARAMETER  selectionMode
        Input Info
    .EXAMPLE
        Input Info
    .OUTPUTS
        Input Info
#>
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $True, ValueFromPipeLine = $True)][PSObject]$PSOTable,
        [Parameter(Mandatory = $False)][String]$SheetName,
        [Parameter(Mandatory = $False)]$WorkbookObject,
        [Parameter(Mandatory = $False)][String]$WorkbookPath,
        [Parameter(Mandatory = $False)][Int32]$MaxColumnWidth = 50
        #[Parameter(Mandatory = $False)][Switch][Boolean]$Save = $False
    )
    
    Add-Type -AssemblyName Microsoft.Office.Interop.Excel

    $excel = New-Object -ComObject excel.application 
    $excel.visible = $False    
    
    # If path is provided then open workbook
    If($WorkbookPath){
        Try{
            $WorkbookObject = $excel.Workbooks.Open($WorkbookPath)
        }
        Catch{}
    }
    
    # If no path and no object is passed to the function then create a workbook
    If(!$WorkbookObject){
        Try{
            $newWorkbook = $True
            $WorkbookObject = $excel.Workbooks.Add() 
            $workbookObject.Worksheets.Item(2).Delete()
            $workbookObject.Worksheets.Item(3).Delete()
        }Catch{
            If(!$WorkbookObject){
                Throw "Unable to create Excel workbook object"
                Break
            }
        }
    }
    
    #Add worksheet and format cells as text
    if(!$newWorkbook){$workSheet = $WorkbookObject.Worksheets.Add()}
    Else{$workSheet = $WorkbookObject.Worksheets.item(1)}
    $workSheet.Cells.NumberFormat = "@"
        
    #Name the worksheet if given
    If($SheetName){
        $origSheetName = $SheetName
        For($i=1; $i -lt 100; $i++){
            Try{
                $workSheet.Name = $SheetName
                break
            }
            Catch{$SheetName = "$origSheetName$i"}
        }
    }
    
    # Set headers
    $headers = $PSOTable | Get-Member | Where-Object{$_.MemberType -eq "NoteProperty"}
    $headers | ForEach-Object{$workSheet.Cells.Item(1,($headers.indexof($_)) + 1) = $_.Name} 
    
    # Format Worksheet Table
    $ListObject = $workSheet.ListObjects.Add([Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange, $WorkSheet.UsedRange, $null ,[Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes)
    $listObject.Name = "Table1"
    $listObject.TableStyle = "TableStyleLight17"
    
    # write collection to worksheet
    $PSOTable | ForEach-Object{
        ForEach($header in $headers){$workSheet.Cells.Item(($PSOTable.indexof($_)) + 2,($headers.indexof($header)) + 1) = ($_."$($header.name)") -join "`n"}
    } 
    
    If($WorkbookPath){
        If($newWorkbook){
            $WorkbookObject.SaveAs($WorkbookPath)
        }Else{
            $WorkbookObject.Save()
        }
    }
    
    $usedRange = $worksheet.UsedRange
    $usedRange.Cells.Font.Size = 9
    $usedrange.WrapText = $True
    $usedRange.HorizontalAlignment = [Microsoft.Office.Interop.Excel.XlHAlign]::xlHAlignLeft.value__
    $usedRange.VerticalAlignment = [Microsoft.Office.Interop.Excel.XlVAlign]::xlVAlignTop.value__
    $usedRange.EntireColumn.ColumnWidth = 70
    $x = $usedRange.EntireColumn.Autofit()
    $x = $usedRange.EntireRow.Autofit()
    $headers | ForEach-Object{$workSheet.Cells.Item(1,($headers.indexof($_)) + 1).Font.Size = 10} 
    

    ForEach($column in $workSheet.UsedRange.EntireColumn){
        If($column.ColumnWidth -gt $MaxColumnWidth){$column.ColumnWidth = $MaxColumnWidth}
    }
    $excel.visible = $True
    
    # Cleanup
    $x = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Remove-Variable excel
}

#$col = Import-Csv ..\Collection.csv
#Export-Excel $col
