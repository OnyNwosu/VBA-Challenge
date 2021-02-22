Attribute VB_Name = "Module1"
Sub MultiYearStockData()

    For Each ws In Worksheets
        Dim TotalStockVolume As Double
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim RowCount As Long
        Dim PercentChange As Double
        Dim SummaryRow As Long
        Dim OpenPrice As Double
        Dim ClosePrice As Double
    
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "YearlyChange"
    Range("K1").Value = "PercentChange"
    Range("L1").Value = "TotalStockVolume"
    
    
    PercentChange = 0
    YearlyChange = 0
    OpenPrice = ws.Cells(2, 3).Value
    ClosePrice = 0
    SummaryRow = 2
    TotalStockVolume = 0
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
     
    
        For i = 2 To RowCount
            If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 12).Value = TotalStockVolume
                ClosePrice = ws.Cells(i, 6).Value
                YearlyChange = ClosePrice - OpenPrice
                ws.Cells(SummaryRow, 10).Value = YearlyChange
            If OpenPrice <> 0 Then
                PercentChange = (ClosePrice - OpenPrice) / OpenPrice
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
                ws.Cells(SummaryRow, 11).Value = PercentChange
                SummaryRow = SummaryRow + 1
                TotalStockVolume = 0
                OpenPrice = ws.Cells(i + 1, 3).Value
                ClosedPrice = 0
                YearlyChange = 0
             End If
        If ws.Cells(SummaryRow, 10).Value > 0 Then
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
            Else
            ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
        End If
        
        End If
        
    Next i
    
    Range("A:L").Columns.AutoFit
    
    Next ws
End Sub




