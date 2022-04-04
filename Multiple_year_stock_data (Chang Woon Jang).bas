Attribute VB_Name = "Module1"
Sub StockCount()

    For Each ws In Worksheets
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Result row variable
        Dim Row As Double
        Row = 2
        
        'open and close value variables
        Dim openval As Double
        Dim closeval As Double
        openval = 0
        closeval = 0
        
        'yearly change and percent change variables
        Dim YearlyChange As Double
        Dim PercentChange As Double
        YearlyChange = 0
        PercentChange = 0
        
        'Total each stock volume variable
        Dim TotalEachStockVol As Double
        TotalEachStockVol = 0
        
        'Result columns in a first row
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        For i = 2 To LastRow
            
            'Accumulate each stock volume
            TotalEachStockVol = TotalEachStockVol + ws.Cells(i, 7).Value
        
            'open value identifier
            If Right(ws.Cells(i, 2).Value, 4) = "0102" Then
                openval = ws.Cells(i, 3).Value
            
            'close value identifier
            ElseIf Right(ws.Cells(i, 2).Value, 4) = "1231" Then
                closeval = ws.Cells(i, 6).Value
                
                'calculations
                YearlyChange = closeval - openval
                PercentChange = (closeval - openval) / openval
            
                ws.Cells(Row, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(Row, 10).Value = YearlyChange
                
                'Background Color by yearly change
                If ws.Cells(Row, 10).Value < 0 Then
                    ws.Cells(Row, 10).Interior.ColorIndex = 3
                
                Else
                    ws.Cells(Row, 10).Interior.ColorIndex = 4
                
                End If
                
                'Insert results in each column
                ws.Cells(Row, 11).Value = PercentChange
                ws.Cells(Row, 11).NumberFormat = "0.00%"
                ws.Cells(Row, 12).Value = TotalEachStockVol
            
                'Row increase for the next ticker
                Row = Row + 1
                'Initiation stock volume to 0 for the next ticker
                TotalEachStockVol = 0
            
            End If
            
        Next i
        
    ws.Range("I:L").Columns.AutoFit
        
    Next ws
    
    'calling a sub statement to summarize the greatest values
    Greatest_Value

End Sub
Sub Greatest_Value()

    For Each ws In Worksheets
    
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        Dim GIncrease As Double
        Dim GDecrease As Double
        Dim GVolume As Double
        
        GIncrease = 0
        GDecrease = 0
        GVolume = 0
        
        Dim GITicker As String
        Dim GDTicker As String
        Dim GVTicker As String
        
        For i = 2 To LastRow
        
            'Find Ticker and Value for Greatest Increase
            If GIncrease < ws.Cells(i, 11).Value Then
                GIncrease = ws.Cells(i, 11).Value
                GITicker = ws.Cells(i, 9).Value
            End If
            
            'Find Ticker and Value for Greatest Decrease
            If GDecrease > ws.Cells(i, 11).Value Then
                GDecrease = ws.Cells(i, 11).Value
                GDTicker = ws.Cells(i, 9).Value
            End If
            
            'Find Ticker and Value for Greatest Volume
            If GVolume < ws.Cells(i, 12).Value Then
                GVolume = ws.Cells(i, 12).Value
                GVTicker = ws.Cells(i, 9).Value
            End If
             
        Next i
    
    'Insert Results
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ws.Cells(2, 16).Value = GITicker
    ws.Cells(2, 17).Value = GIncrease
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    ws.Cells(3, 16).Value = GDTicker
    ws.Cells(3, 17).Value = GDecrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ws.Cells(4, 16).Value = GVTicker
    ws.Cells(4, 17).Value = GVolume
    
    'Column Autofit on the result sections
    ws.Range("O:Q").Columns.AutoFit
    
    Next ws
    
End Sub
