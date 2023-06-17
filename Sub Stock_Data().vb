Sub Stock_Data()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        'declare variables
        Dim data_count As Long
        Dim array_counter As Integer
        
        array_counter = 1
        'define number of data points in each worksheet to set uppper bound for loop
        data_count = WorksheetFunction.Count(ws.Columns("C"))
        'defines the number of unique stock tickers
        For i = 2 To data_count
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                 array_counter = array_counter + 1
                 
            End If
        Next i
        
        'define array variables to store unique values for each ticker
        Dim tickers()
        Dim opening_price()
        Dim closing_price()
        Dim yearly_change()
        Dim stock_volume()
        ReDim tickers(array_counter)
        ReDim opening_price(array_counter)
        ReDim closing_price(array_counter)
        ReDim yearly_change(array_counter)
        ReDim stock_volume(array_counter)
        
        Dim placeholder As Integer
        tickers(0) = ws.Cells(2, 1)
        opening_price(0) = ws.Cells(2, 3)
        placeholder = 0
        For i = 2 To data_count
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                 tickers(placeholder + 1) = ws.Cells(i, 1)
                 stock_volume(placeholder) = stock_volume(placeholder) + ws.Cells(i, 7)
                 closing_price(placeholder) = ws.Cells(i, 6)
                 opening_price(placeholder + 1) = ws.Cells(i + 1, 3)
                 placeholder = placeholder + 1
            Else
                stock_volume(placeholder) = stock_volume(placeholder) + ws.Cells(i, 7)
            End If
        Next i
        
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
      
        For i = 1 To array_counter
            ws.Cells(i + 1, 9) = tickers(i - 1)
            ws.Cells(i + 1, 10) = closing_price(i - 1) - opening_price(i - 1)
            If (ws.Cells(i + 1, 10) < 0) Then
                ws.Cells(i + 1, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i + 1, 10).Interior.ColorIndex = 4
            End If
            ws.Cells(i + 1, 11) = (closing_price(i - 1) - opening_price(i - 1)) / opening_price(i - 1)
            ws.Cells(i + 1, 12) = stock_volume(i - 1)
        Next i
    Next
End Sub

