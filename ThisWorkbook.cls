
Sub CalculateColumnSums()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastRow2 As Long
    Dim lastColumn As Long
    Dim summarySheet As Worksheet
    Dim totalSum As Double
    Dim countSum As Double
    Dim totalSumF As Double
    Dim countNotZero As Integer
    Dim summarySheetName As String
    Dim filteredRange As Range

    ' Get the last sheet in the workbook
    Set ws = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)

    ' Find the last row and column with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
     
    ' Create a new sheet for the summary
    summarySheetName = "Summary"
    
    ' Find target sheet
    Set summarySheet = ActiveWorkbook.Sheets(summarySheetName)
    
    ' Find the last row and column with data
    lastRow2 = summarySheet.Cells(summarySheet.Rows.Count, 1).End(xlUp).Row

    ' Add headers to the summary sheet
    summarySheet.Cells(1, 1).Value = "Column"
    summarySheet.Cells(1, 2).Value = "Sum"
    
    'fEILD 1 STULPELI IMU
    ' Filter column E for values greater than 10
    ws.Range("E2:E" & lastRow).AutoFilter Field:=1, Criteria1:=">0.10"
    
    ' Get the filtered range
    Set filteredRange = ws.Range("B2:B" & lastRow).SpecialCells(xlCellTypeVisible)
    ' Get the filtered range
    Set filteredRange2 = ws.Range("E2:E" & lastRow).SpecialCells(xlCellTypeVisible)

    ' Calculate and add column sums from columns E to I
    Dim i As Long

    For i = 5 To 9 ' Columns E to I
        'summarySheet.Cells(i - 4 + 2, 1).Value = ws.Cells(1, i).Value
        'summarySheet.Cells(i - 4 + 2, 2).Value = Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i)))
        totalSum = totalSum + Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i)))
        countSum = countSum + Application.WorksheetFunction.CountA(ws.Range(ws.Cells(2, i), ws.Cells(lastRow, i)))
        totalSumF = totalSum + Application.WorksheetFunction.Sum(ws.Range(ws.Cells(2, 6), ws.Cells(lastRow, 6)))
        countNotZero = Application.WorksheetFunction.CountIf(ws.Range(ws.Cells(2, 6), ws.Cells(lastRow, 6)), ">0.00")
    Next i
    ' Write the total sum and count
    summarySheet.Cells(lastRow2 + 1, 1).Value = "Total"
    summarySheet.Cells(lastRow2 + 1, 2).Value = totalSum
    summarySheet.Cells(lastRow2 + 1, 3).Value = countSum
    summarySheet.Cells(lastRow2 + 2, 1).Value = "TotalF"
    summarySheet.Cells(lastRow2 + 2, 2).Value = totalSumF
    summarySheet.Cells(lastRow2 + 3, 1).Value = "COUNT GRATER THAN ZERO"
    summarySheet.Cells(lastRow2 + 3, 2).Value = countNotZero
    
      ' Write the values greater than 10 to a new column
    summarySheet.Cells(lastRow2 + 3, 5).Value = "Values Greater Than 10"
    Dim j As Long
    For j = 1 To filteredRange.Cells.Count
        summarySheet.Cells(lastRow2 + 3 + j, 5).Value = filteredRange.Cells(j).Value
        summarySheet.Cells(lastRow2 + 3 + j, 6).Value = filteredRange2.Cells(j).Value
    Next j
    


End Sub
