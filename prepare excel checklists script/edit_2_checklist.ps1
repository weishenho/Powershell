$jsonFile = get-item "first enter info here.JSON"
$json = Get-Content $jsonFile | Out-String | ConvertFrom-Json

# Specify the path to the Excel file and the WorkSheet Name
$FilePath = get-item $json.checklist2

# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $false

# Open the Excel file and save it in $WorkBook
$WorkBook = $objExcel.Workbooks.Open($FilePath)

## Cover Sheet
$curSheet = $WorkBook.Worksheets.Item("Cover Sheet")

$jiraTicket = $json.coverSheet.jiraTicket
$coversheetSystem = $json.coverSheet.system
$module = $json.coverSheet.module
$funcArea = $json.coverSheet.functionalArea
$techDesigner = $json.coverSheet.technicalDesigner
$devName = $json.coverSheet.developerName
$peerName = $json.coverSheet.peerReviewer
$dateToday = Get-Date -format "dd-MMM-yyyy"

#Jira Ticket
$curSheet.Cells.Item(7, 3) = $jiraTicket
$link = "https://jira/" + $jiraTicket
$curSheet.Hyperlinks.Add($curSheet.Cells.Item(7, 3), $link)

#System
$curSheet.Cells.Item(8, 3) = $coversheetSystem
#System
$curSheet.Cells.Item(9, 3) = $module
#Functional Area Ticket
$curSheet.Cells.Item(10, 3) = $funcArea
#checklist prepare date
$curSheet.Cells.Item(11, 3) = $dateToday
#Technical Designer
$curSheet.Cells.Item(12, 3) = $techDesigner
#Developer Name
$curSheet.Cells.Item(13, 3) = $devName
#Peer Reviewer
$curSheet.Cells.Item(14, 3) = $peerName
#Reviewed Date
$curSheet.Cells.Item(15, 3) = $dateToday


## x Development Sheet
$xItemsReviewed = $json.xDevelopment.xItemsReviewed -join "`n"
$curSheet = $WorkBook.worksheets.item('x Development')
$curSheet.Range("C3:G5") = $xItemsReviewed

## y Development Sheet
$yReviewed = $json.yDevelopment.yReviewed -join "`n"
$curSheet = $WorkBook.worksheets.item('y Development')
$curSheet.Range("C3:G5") = $yReviewed

## z Development Sheet
$zReviewed = $json.zDevelopment.zReviewed -join "`n"
$curSheet = $WorkBook.worksheets.item('z Development')
$curSheet.Range("C3:G5") = $zReviewed

#Save File
$peerReviewFileName = "Checklist_" + $jiraTicket + ".xlsx"

$objExcel.DisplayAlerts = $false
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$currLocation = (Resolve-Path .\).Path
$objExcel.ActiveWorkbook.SaveAs($currLocation + "\" + $peerReviewFileName, $xlFixedFormat)

# Quit Editing Peer Review
$objExcel.ActiveWorkbook.Close()
$objExcel.Quit()

[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook)
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel)

Remove-Variable -Name objExcel