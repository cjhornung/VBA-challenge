Sub Stock_Data()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        'declare variables to identify the amount of data entries and the amounts of unique ticker entries
        Dim data_count As Long
        Dim array_counter As Integer
        
        array_counter = 1
        'define number of data points in each worksheet to set uppper bound for loop needed to define number of unique ticker entries
        data_count = WorksheetFunction.Count(ws.Columns("C"))
        'define number of unique ticker entries
        For i = 2 To data_count
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                 array_counter = array_counter + 1
                 
            End If
        Next i
        
        'define dynamic array variables to store unique values for each ticker
        Dim tickers()
        Dim opening_price()
        Dim closing_price()
        Dim yearly_change()
        Dim stock_volume()
        ReDim tickers(array_counter - 1)
        ReDim opening_price(array_counter - 1)
        ReDim closing_price(array_counter - 1)
        ReDim yearly_change(array_counter - 1)
        ReDim stock_volume(array_counter - 1)
        Dim placeholder As Integer
        'set intial array values for variables that require it
        tickers(0) = ws.Cells(2, 1)
        opening_price(0) = ws.Cells(2, 3)
        placeholder = 0
        'loop to set array values for each unique ticker
        For i = 2 To data_count
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                 tickers(placeholder + 1) = ws.Cells(i + 1, 1)
                 stock_volume(placeholder) = stock_volume(placeholder) + ws.Cells(i, 7)
                 closing_price(placeholder) = ws.Cells(i, 6)
                 opening_price(placeholder + 1) = ws.Cells(i + 1, 3)
                 placeholder = placeholder + 1
            Else
                stock_volume(placeholder) = stock_volume(placeholder) + ws.Cells(i, 7)
            End If
        Next i
        'set final array values for variables that require it
        closing_price(array_counter - 1) = ws.Cells(data_count + 1, 6)
        'Column Headers for Output Table
        ws.Cells(1, 9) = "Ticker"
        ws.Cells(1, 10) = "Yearly Change"
        ws.Cells(1, 11) = "Percent Change"
        ws.Cells(1, 12) = "Total Stock Volume"
        
        'Declare variables for Summary Table
        Dim greatest_percent_increase As Double
        Dim greatest_percent_decrease As Double
        Dim greatest_total_volume As Long
        greatest_percent_increase = 0
        greatest_percent_decrease = 0
        greatest_total_volume = 0
        Dim ticker_greatest_percent_increase As String
        Dim ticker_greatest_percent_decrease As String
        Dim ticker_greatest_total_volume As String
 
        'Loop to output findings to Output Table
        For i = 1 To array_counter
            ws.Cells(i + 1, 9) = tickers(i - 1)
            ws.Cells(i + 1, 10) = closing_price(i - 1) - opening_price(i - 1)
            If (ws.Cells(i + 1, 10) < 0) Then
                ws.Cells(i + 1, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(i + 1, 10).Interior.ColorIndex = 4
            End If
            ws.Cells(i + 1, 11) = (closing_price(i - 1) - opening_price(i - 1)) / opening_price(i - 1)
            ws.Cells(i + 1, 11).NumberFormat = "0.00%"
            If (ws.Cells(i + 1, 11) < 0) Then
                ws.Cells(i + 1, 11).Interior.ColorIndex = 3
            Else
                ws.Cells(i + 1, 11).Interior.ColorIndex = 4
            End If
            ws.Cells(i + 1, 12) = stock_volume(i - 1)
            'Update Summary Table Values
            If ws.Cells(i + 1, 11) > greatest_percent_increase Then
                ticker_greatest_percent_increase = ws.Cells(i + 1, 9)
                greatest_percent_increase = ws.Cells(i + 1, 11)
            End If
            If ws.Cells(i + 1, 11) < greatest_percent_decrease Then
                ticker_greatest_percent_decrease = ws.Cells(i + 1, 9)
                greatest_percent_decrease = ws.Cells(i + 1, 11)
            End If
             If ws.Cells(i + 1, 12) > greatest_total_volume Then
                ticker_greatest_total_volume = ws.Cells(i + 1, 9)
                greatest_total_volume = ws.Cells(i + 1, 12)
            End If
        Next i
        'Setup Summary Table
        ws.Cells(2, 15) = "Greatest % increase"
        ws.Cells(2, 16) = ticker_greatest_percent_increase
        ws.Cells(2, 17) = greatest_percent_increase
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 15) = "Greatest % decrease"
        ws.Cells(3, 16) = ticker_greatest_percent_decrease
        ws.Cells(3, 17) = greatest_percent_decrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 15) = "Greatest total volume"
        ws.Cells(4, 16) = ticker_greatest_total_volume
        ws.Cells(4, 17) = greatest_total_volume
        
        ws.Cells(1, 16) = "Ticker"
        ws.Cells(1, 17) = "Value"
        
        
        'AutoFit Formatting for display purposes
        ws.Columns("A:Q").AutoFit
        
    Next
End Sub

