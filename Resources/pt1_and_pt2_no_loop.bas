Attribute VB_Name = "Module2"
Sub Output_Stock_Data()
    ' Part 1 & 2 before adding the loop through all worksheets
    
    'Part 1 Exercise
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Yearly Change"
    Cells(1, "K").Value = "Percent Change"
    Cells(1, "L").Value = "Total Stock Volume"
    
    Dim i, row_counter As Integer
    Dim lastrow, total_stock_volume, open_val, close_val As Double
    
    row_counter = 2
    open_val = Cells(2, "C").Value
    close_val = Cells(2, "F").Value
    total_stock_volume = Cells(2, "G").Value
    lastrow = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    For i = 3 To lastrow
        ' if under the same ticker, update close_val
        If Cells(i, "A").Value = Cells(i - 1, "A").Value Then
            close_val = Cells(i, "F").Value
            total_stock_volume = total_stock_volume + Cells(i, "G").Value
        ' if the next ticker value is a different one, calculate the yearly for that ticker then update open_val and close_val
        Else
            Cells(row_counter, "I").Value = Cells(i - 1, "A").Value
            Cells(row_counter, "J").Value = Format(close_val - open_val, "#.00")
            If Cells(row_counter, "J") >= 0 Then
                Cells(row_counter, "J").Interior.ColorIndex = 4
            Else
                Cells(row_counter, "J").Interior.ColorIndex = 3
            End If
            Cells(row_counter, "K").Value = Format((Cells(row_counter, "J").Value) / open_val, "###0.00%")
            Cells(row_counter, "L").Value = total_stock_volume
            open_val = Cells(i, "C").Value
            close_val = Cells(i + 1, "F").Value
            total_stock_volume = Cells(i, "G").Value
            row_counter = row_counter + 1
            
        End If
    Next i
    
    ' Part 2 Exercise
    Cells(2, "O").Value = "Greatest % Increase"
    Cells(3, "O").Value = "Greatest % Decrease"
    Cells(4, "O").Value = "Greatest Total Volume"
    Cells(1, "P").Value = "Ticker"
    Cells(1, "Q").Value = "Value"
    
    Dim inc_ticker, dec_ticker, vol_ticker As String
    Dim greatest_inc, greatest_dec, greatest_vol As Double
    
    ' assigning greatest % increase to first percent change val, will update as we find a larger value
    inc_ticker = Cells(2, "I").Value
    greatest_inc = Cells(2, "K").Value
    ' assigning greatest % decrease to first percent change val, will update as we find a larger negative value
    dec_ticker = Cells(2, "I").Value
    greatest_dec = Cells(2, "K").Value
    ' assigning greatest total stock volume to first val, will update as we find a larger total volume
    vol_ticker = Cells(2, "I").Value
    greatest_vol = Cells(2, "L").Value
    
    For i = 3 To lastrow
        ' Updating Greatest % Increase: Check if row has larger % Change than what is saved
        If Cells(i, "K").Value > greatest_inc Then
            inc_ticker = Cells(i, "I").Value
            greatest_inc = Cells(i, "K").Value
        End If
        
         ' Updating Greatest % Decrease: Check if row has larger (negative) % Change than what is saved
        If Cells(i, "K").Value < greatest_dec Then
            dec_ticker = Cells(i, "I").Value
            greatest_dec = Cells(i, "K").Value
        End If
        
        ' Updating Greatest Total Stock Volume: Check if row has larger total stock volume than what is saved
        If Cells(i, "L").Value > greatest_vol Then
            vol_ticker = Cells(i, "I").Value
            greatest_vol = Cells(i, "L").Value
        End If
    Next i
    
    Cells(2, "P").Value = inc_ticker
    Cells(2, "Q").Value = Format(greatest_inc, "#.00%")
    Cells(3, "P").Value = dec_ticker
    Cells(3, "Q").Value = Format(greatest_dec, "#.00%")
    Cells(4, "P").Value = vol_ticker
    Cells(4, "Q").Value = greatest_vol
End Sub
