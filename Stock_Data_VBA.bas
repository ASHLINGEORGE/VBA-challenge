Attribute VB_Name = "Stock_Data"
Sub StockData()
    ' Declare variables
    Dim ws As Worksheet
    Dim row As Long
    Dim LastRow As Long
    Dim currentticker As String
    Dim preticker As String
    Dim postticker As String
    Dim openvalue As Double
    Dim closevalue As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalvolume As Double
    Dim percentincrease As Double
    Dim percentdecrease As Double
    Dim greatestvolume As Double
    Dim increaseticker As String
    Dim decreaseticker As String
    Dim greatestticker As String
    
    ' Loop through each worksheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Set column headers for data and results
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Initialize variables
        row = 2
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        ' Loop through rows of data
        For i = 2 To LastRow
            currentticker = ws.Cells(i, 1).Value
            preticker = ws.Cells(i - 1, 1).Value
            postticker = ws.Cells(i + 1, 1).Value
            
            ' Calculate open value for each new ticker
            If preticker <> currentticker Then
                openvalue = ws.Cells(i, 3).Value
            End If
            
            ' Accumulate the total stock volume
            totalvolume = totalvolume + ws.Cells(i, 7).Value
            
            ' When the ticker changes, calculate and record data
            If currentticker <> postticker Then
                closevalue = ws.Cells(i, 6).Value
                yearlychange = closevalue - openvalue
                If openvalue <> 0 Then
                    result = (yearlychange / openvalue) * 100
                    percentchange = Round(result, 2)
                Else
                    percentchange = 0
                End If
                
                ' Record data in the worksheet
                ws.Cells(row, 9).Value = currentticker
                ws.Cells(row, 10).Value = yearlychange
                ws.Cells(row, 11).Value = CStr(percentchange) + "%"
                ws.Cells(row, 12).Value = totalvolume
            
            
                ' Color cells based on yearly change
                If yearlychange < 0 Then
                    ws.Cells(row, 10).Interior.Color = RGB(255, 0, 0) ' Red
                    ws.Cells(row, 11).Interior.Color = RGB(255, 0, 0) ' Red
                Else
                    ws.Cells(row, 10).Interior.Color = RGB(0, 255, 0) ' Green
                    ws.Cells(row, 11).Interior.Color = RGB(0, 255, 0) ' Green
                End If
                
                ' Update variables for greatest values
                If percentincrease < percentchange Then
                    percentincrease = percentchange
                    increaseticker = currentticker
                End If
                
                If percentdecrease > percentchange Then
                    percentdecrease = percentchange
                    decreaseticker = currentticker
                End If
                
                If greatestvolume < totalvolume Then
                    greatestvolume = totalvolume
                    greatestticker = currentticker
                End If
            
                ' Reset variables for the next row
                openvalue = 0
                closevalue = 0
                totalvolume = 0
                row = row + 1
            End If
            ' Move to the next row
        Next i
        
        ' Record the results in the worksheet
        ws.Cells(2, 16).Value = increaseticker
        ws.Cells(3, 16).Value = decreaseticker
        ws.Cells(4, 16).Value = greatestticker
        ws.Cells(2, 17).Value = CStr(percentincrease) + "%"
        ws.Cells(3, 17).Value = CStr(percentdecrease) + "%"
        ws.Cells(4, 17).Value = greatestvolume
        
        ' Reset variables for the next worksheet
        percentincrease = 0
        percentdecrease = 0
        greatestvolume = 0
    Next ws
End Sub

