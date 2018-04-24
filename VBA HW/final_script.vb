Sub stock()
    Dim open_price, close_price, yearly_change As Double
    Dim ticker_start_row, ticker_end_row As Double
    Dim result_row As Integer
    Dim ticker_name As String
    Dim total_vol As Double
    
    For Each ws In Worksheets
        ' Initialize
        ticker_start_row = 2
        ticker_end_row = 0
        result_row = 2
        total_vol = 0
        yearl_change = 0
        
        
        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Loop through all rows
        For Row = 2 To lastRow
        
            If ws.Cells(Row, 1) <> ws.Cells(Row + 1, 1) Then
                ' A new ticker starts
                ticker_name = ws.Cells(Row, 1)
                
                ticker_end_row = Row
                
                ' Calculate the total stock volume of the same ticker
                For ticker_row = ticker_start_row To ticker_end_row
                    total_vol = total_vol + ws.Cells(ticker_row, 7).Value
                Next ticker_row
                
                ' Calculate yearly change
                open_price = ws.Cells(ticker_start_row, 3).Value
                close_price = ws.Cells(ticker_end_row, 6).Value
                yearly_change = close_price - open_price
                 
                
                ' ticker name
                ws.Cells(result_row, 9).Value = ticker_name
                
                ' yearly change
                ws.Cells(result_row, 10).Value = yearly_change
                
                ' percentage change
                If open_price = 0 Then
                    ws.Cells(result_row, 11).Value = 0
                Else
                    ws.Cells(result_row, 11).Value = yearly_change / open_price
                End If
                
                ' total volume
                ws.Cells(result_row, 12).Value = total_vol
                
                
                result_row = result_row + 1
                ticker_start_row = Row + 1
                total_vol = 0
                yearly_change = 0
            
            End If
        Next Row
        
        ' Loop through the percent changes
        Dim last_row_percent_change As Double
        last_row_percent_change = ws.Cells(Rows.Count, 9).End(xlUp).Row
        Dim min, max, temp_value, max_total_vol, temp_vol As Double
        Dim min_ticker, max_ticker, max_vol_ticker, temp_ticker As String
        
        min = 0
        max = 0
        max_total_vol = 0
        
        For percent_change_row = 2 To last_row_percent_change
            temp_value = ws.Cells(percent_change_row, 11).Value
            temp_ticker = ws.Cells(percent_change_row, 9).Value
            temp_vol = ws.Cells(percent_change_row, 12).Value
            
            If temp_value > max Then
                max = temp_value
                max_ticker = temp_ticker
            End If
            
            If temp_value < min Then
                min = temp_value
                min_ticker = temp_ticker
            End If
            
            If temp_vol > max_total_vol Then
                max_total_vol = temp_vol
                max_vol_ticker = temp_ticker
            End If
            
        Next percent_change_row
        
        ' Write the greatest increase/decrease to cells
        ws.Cells(2, 16).Value = max_ticker
        ws.Cells(2, 17).Value = max
        
        ws.Cells(3, 16).Value = min_ticker
        ws.Cells(3, 17).Value = min
        
        ws.Cells(4, 16).Value = max_vol_ticker
        ws.Cells(4, 17).Value = max_total_vol
        
        
    ' The next worksheet
    Next
    
End Sub



