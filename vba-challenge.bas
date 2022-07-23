Attribute VB_Name = "Module1"
Sub stocks_info()

For Each ws In Worksheets
    Dim t_ticker, num_tickers, num_rows As Long
    Dim sum_vol As LongPtr
    
    'Set up initial row of headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly_Change"
    ws.Range("K1").Value = "Percent_Change"
    ws.Range("L1").Value = "Total Volume"
    ws.Range("N1").Value = "Greatest % Increase"
    ws.Range("N2").Value = "Greatest % Decrease"
    ws.Range("N3").Value = "Greatest Total Volume"
    
    'Find the number of rows and columns in the worksheet
    num_rows = ws.Cells(Rows.Count, 1).End(xlUp).Row
    num_col = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    't_row is the summary row for the table
    t_row = 2
    
    'Initialize the beginning of the year openning price and the sum of the stock volume
    o_price = ws.Cells(2, 3).Value
    sum_vol = 0
    
    For r_ticker = 2 To num_rows
        'check if previous ticker is different from current ticker
        If ws.Cells(r_ticker, 1).Value <> ws.Cells(r_ticker + 1, 1).Value Then
            'Yearly Change
            ws.Cells(t_row, 10).Value = ws.Cells(r_ticker, 6).Value - o_price
            ws.Cells(t_row, 11).Value = ws.Cells(t_row, 10).Value / o_price
            ws.Cells(t_row, 11).Style = "Percent"
            ws.Cells(t_row, 10).Style = "Currency"
            o_price = ws.Cells(r_ticker + 1, 3).Value
            
            'Total Stock Volume
            ws.Cells(t_row, 12).Value = sum_vol
            sum_vol = 0
            'Add new ticker to the summary table
            ws.Cells(t_row, 9).Value = ws.Cells(r_ticker, 1).Value
            
            
            'Color Formatting
            If ws.Cells(t_row, 10).Value > 0 Then
                ws.Cells(t_row, 10).Interior.ColorIndex = 4 '4 is green
                ws.Cells(t_row, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(t_row, 10).Interior.ColorIndex = 3 '3 is red
                ws.Cells(t_row, 11).Interior.ColorIndex = 3
            End If
            
            'Increase t_row such that new tickers can be added to the table
            t_row = t_row + 1
        Else
           sum_vol = sum_vol + ws.Cells(r_ticker, 7).Value
        End If
    Next r_ticker
    
    
    'Bonus Problem: Greatest increase, decrease, and volume
    'g_incr_row is the row with the greatest increase. g_decr_row is the row with the greatest decrease. g_tot_v is the greatest total volume
    g_incr_row = 2
    g_decr_row = 2
    g_tot_v = 2
    
    num_tickers = ws.Cells(Rows.Count, 10).End(xlUp).Row
    For r = 3 To num_tickers
        'Greatest increase
        If ws.Cells(r, 10).Value > ws.Cells(g_incr_row, 10).Value Then
            g_incr_row = r
        End If
        
        'Greatest decrease
        If ws.Cells(r, 10).Value < ws.Cells(g_decr_row, 10).Value Then
            g_decr_row = r
        End If
        
        'Greatest Vol
        If ws.Cells(r, 12).Value > ws.Cells(g_tot_v, 12).Value Then
            g_tot_v = r
        End If
    Next r
    
    ws.Range("O1").Value = ws.Range("I" & g_incr_row).Value
    ws.Range("O2").Value = ws.Range("I" & g_decr_row).Value
    ws.Range("O3").Value = ws.Range("I" & g_tot_v).Value

    ws.Range("P1").Value = ws.Range("K" & g_incr_row).Value
    ws.Range("P2").Value = ws.Range("K" & g_decr_row).Value
    ws.Range("P3").Value = ws.Range("L" & g_tot_v).Value
    ws.Range("P1:P2").Style = "Percent"
    
    'Formatting columns to fit
    ws.Columns("I:O").AutoFit
Next ws
End Sub
