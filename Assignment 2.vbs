Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim ticker As String
    Dim openPrice As Double
    Dim closePrice As Double
    Dim volume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim outputRow As Long


    For Each ws In ThisWorkbook.Worksheets
        outputRow = 2
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    

        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            openPrice = ws.Cells(i, 3).Value
            closePrice = ws.Cells(i, 6).Value
            volume = ws.Cells(i, 7).Value


            quarterlyChange = closePrice - openPrice
            percentChange = (quarterlyChange / openPrice) * 100

 
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = quarterlyChange
            ws.Cells(outputRow, 11).Value = percentChange
            ws.Cells(outputRow, 12).Value = volume
            
            outputRow = outputRow + 1
        Next i
    Next ws

    With ThisWorkbook.Worksheets(1).Range("J2:J" & outputRow - 1)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
        .FormatConditions(1).Interior.Color = RGB(0, 255, 0)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        .FormatConditions(2).Interior.Color = RGB(255, 0, 0)
    End With

    MsgBox "Analysis complete!"
End Sub
