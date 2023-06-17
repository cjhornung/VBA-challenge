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
        Dim stock_volume()
        ReDim tickers(array_counter)
        ReDim opening_price(array_counter)
        ReDim closing_price(array_counter)
        ReDim stock_volume(array_counter)
        
        Dim placeholder As Integer
        tickers(0) = ws.Cells(2, 1)
        pening_price(0) = ws.Cells(2, 3)
        stock_volume(0) = ws.Cells(2, 7)
        placeholder = 1
        For i = 2 To data_count
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                 tickers(placeholder) = ws.Cells(i, 1)
                 placeholder = placeholder + 1
            End If
        Next i
        
        
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
      
        For i = 1 To array_counter
            ws.Cells(i + 1, 9) = tickers(i - 1)
            
        Next i
    Next
End Sub

