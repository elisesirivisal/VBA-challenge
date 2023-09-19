Attribute VB_Name = "Module1"
Sub Output_Stock_Data()
    Dim i, j, row_counter As Integer
    Dim lastrow, total_stock_volume, open_val, close_val, greatest_inc, greatest_dec, greatest_vol As Double
    Dim inc_ticker, dec_ticker, vol_ticker As String
    
    'get number of worksheets to traverse through
    num_wkshts = ActiveWorkbook.Worksheets.Count
    
    For j = 1 To num_wkshts
        'Part 1 Exercise
        ActiveWorkbook.Worksheets(j).Cells(1, "I").Value = "Ticker"
        ActiveWorkbook.Worksheets(j).Cells(1, "J").Value = "Yearly Change"
        ActiveWorkbook.Worksheets(j).Cells(1, "K").Value = "Percent Change"
        ActiveWorkbook.Worksheets(j).Cells(1, "L").Value = "Total Stock Volume"
        
        row_counter = 2
        open_val = ActiveWorkbook.Worksheets(j).Cells(2, "C").Value
        close_val = ActiveWorkbook.Worksheets(j).Cells(2, "F").Value
        total_stock_volume = ActiveWorkbook.Worksheets(j).Cells(2, "G").Value
        lastrow = ActiveWorkbook.Worksheets(j).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        
        For i = 3 To lastrow
            ' if under the same ticker, update close_val
            If ActiveWorkbook.Worksheets(j).Cells(i, "A").Value = ActiveWorkbook.Worksheets(j).Cells(i - 1, "A").Value Then
                close_val = ActiveWorkbook.Worksheets(j).Cells(i, "F").Value
                total_stock_volume = total_stock_volume + ActiveWorkbook.Worksheets(j).Cells(i, "G").Value
            ' if the next ticker value is a different one, calculate the yearly for that ticker then update open_val and close_val
            Else
                ActiveWorkbook.Worksheets(j).Cells(row_counter, "I").Value = ActiveWorkbook.Worksheets(j).Cells(i - 1, "A").Value
                ActiveWorkbook.Worksheets(j).Cells(row_counter, "J").Value = Format(close_val - open_val, "#.00")
                If ActiveWorkbook.Worksheets(j).Cells(row_counter, "J") >= 0 Then
                    ActiveWorkbook.Worksheets(j).Cells(row_counter, "J").Interior.ColorIndex = 4
                Else
                    ActiveWorkbook.Worksheets(j).Cells(row_counter, "J").Interior.ColorIndex = 3
                End If
            ActiveWorkbook.Worksheets(j).Cells(row_counter, "K").Value = Format((ActiveWorkbook.Worksheets(j).Cells(row_counter, "J").Value) / open_val, "###0.00%")
            ActiveWorkbook.Worksheets(j).Cells(row_counter, "L").Value = total_stock_volume
            open_val = ActiveWorkbook.Worksheets(j).Cells(i, "C").Value
            close_val = ActiveWorkbook.Worksheets(j).Cells(i + 1, "F").Value
            total_stock_volume = ActiveWorkbook.Worksheets(j).Cells(i, "G").Value
            row_counter = row_counter + 1
                
            End If
        Next i
        
        ' Part 2 Exercise
        ActiveWorkbook.Worksheets(j).Cells(2, "O").Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(j).Cells(3, "O").Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(j).Cells(4, "O").Value = "Greatest Total Volume"
        ActiveWorkbook.Worksheets(j).Cells(1, "P").Value = "Ticker"
        ActiveWorkbook.Worksheets(j).Cells(1, "Q").Value = "Value"
        
        ' assigning greatest % increase to first percent change val, will update as we find a larger value
        inc_ticker = ActiveWorkbook.Worksheets(j).Cells(2, "I").Value
        greatest_inc = ActiveWorkbook.Worksheets(j).Cells(2, "K").Value
        ' assigning greatest % decrease to first percent change val, will update as we find a larger negative value
        dec_ticker = ActiveWorkbook.Worksheets(j).Cells(2, "I").Value
        greatest_dec = ActiveWorkbook.Worksheets(j).Cells(2, "K").Value
        ' assigning greatest total stock volume to first val, will update as we find a larger total volume
        vol_ticker = ActiveWorkbook.Worksheets(j).Cells(2, "I").Value
        greatest_vol = ActiveWorkbook.Worksheets(j).Cells(2, "L").Value
        
        For i = 3 To lastrow
            ' Updating Greatest % Increase: Check if row has larger % Change than what is saved
            If ActiveWorkbook.Worksheets(j).Cells(i, "K").Value > greatest_inc Then
                inc_ticker = ActiveWorkbook.Worksheets(j).Cells(i, "I").Value
                greatest_inc = ActiveWorkbook.Worksheets(j).Cells(i, "K").Value
            End If
            
             ' Updating Greatest % Decrease: Check if row has larger (negative) % Change than what is saved
            If ActiveWorkbook.Worksheets(j).Cells(i, "K").Value < greatest_dec Then
                dec_ticker = ActiveWorkbook.Worksheets(j).Cells(i, "I").Value
                greatest_dec = ActiveWorkbook.Worksheets(j).Cells(i, "K").Value
            End If
            
            ' Updating Greatest Total Stock Volume: Check if row has larger total stock volume than what is saved
            If ActiveWorkbook.Worksheets(j).Cells(i, "L").Value > greatest_vol Then
                vol_ticker = ActiveWorkbook.Worksheets(j).Cells(i, "I").Value
                greatest_vol = ActiveWorkbook.Worksheets(j).Cells(i, "L").Value
            End If
        Next i
        
        ActiveWorkbook.Worksheets(j).Cells(2, "P").Value = inc_ticker
        ActiveWorkbook.Worksheets(j).Cells(2, "Q").Value = Format(greatest_inc, "#.00%")
        ActiveWorkbook.Worksheets(j).Cells(3, "P").Value = dec_ticker
        ActiveWorkbook.Worksheets(j).Cells(3, "Q").Value = Format(greatest_dec, "#.00%")
        ActiveWorkbook.Worksheets(j).Cells(4, "P").Value = vol_ticker
        ActiveWorkbook.Worksheets(j).Cells(4, "Q").Value = greatest_vol
    Next j
End Sub

