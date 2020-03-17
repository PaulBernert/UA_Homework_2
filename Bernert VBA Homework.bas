Attribute VB_Name = "bernert_hw_module"
'Initialize module
Sub homework_()

    'Iterate over worksheets
    For Each ws In Worksheets

        'Add labels and misc. text
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Year Open Price"
        ws.Range("K1").Value = "Year Close Price"
        ws.Range("L1").Value = "Change Price Per share"
        ws.Range("M1").Value = "Change percent"
        ws.Range("N1").Value = "Total Volume"

        'Create dimensions
        Dim year_open, year_close, soln_sub, soln_pct, max_val, min_val As Double
        Dim ticker, min_ticker, max_ticker, max_vol_ticker As String
        Dim summary_row, row_len, row_counter As Integer
        Dim vol, max_vol As LongLong
    
        'Assign default values
        summary_row = 2
        row_count = 1
        row_len = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Loop through all tickers
        For i = 2 To row_len
            
            'Check if we are still within the same ticker name, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
                'Assign values
                ticker = ws.Cells(i, 1).Value
                year_close = ws.Cells(i, 6).Value

                'If open has a zero..
                If (year_open = 0) Then
                    soln_sub = 0
                    soln_pct = 0
                'If open does not have a zero..
                Else
                    'Calculate soln_sub and soln_pct
                    soln_sub = year_close - year_open
                    soln_pct = (year_close / year_open) - 1
                End If
                
                'Populate values
                ws.Range("I" & summary_row).Value = ticker
                ws.Range("J" & summary_row).Value = year_open
                ws.Range("K" & summary_row).Value = year_close
                ws.Range("N" & summary_row).Value = vol
                ws.Range("L" & summary_row).Value = soln_sub
                ws.Range("M" & summary_row).Value = soln_pct
                
                'Formatting
                ws.Range("L" & summary_row).NumberFormat = "$#.##"
                ws.Range("M" & summary_row).NumberFormat = "0.0%"

                'Conditional formatting
                If ws.Range("L" & summary_row).Value < 0 Then
                    ws.Range("L" & summary_row).Interior.ColorIndex = 3
                Else
                    ws.Range("L" & summary_row).Interior.ColorIndex = 10
                End If
                
                'Checking for max_val and min_val
                If ws.Range("M" & summary_row).Value > max_val Then
                    max_val = ws.Range("M" & summary_row).Value
                    max_tic = ws.Range("I" & summary_row).Value
                ElseIf ws.Range("M" & summary_row).Value < min_val Then
                    min_val = ws.Range("M" & summary_row).Value
                    min_tic = ws.Range("I" & summary_row).Value
                End If
            
                'Reset counters and values
                summary_row = summary_row + 1
                vol = 0
                row_count = 1
            
            'Check if we are still within the same ticker name, if it is...
            Else
                
                'Assign values
                vol = vol + ws.Cells(i, 7).Value

                'Checking for max_vol
                If vol > max_vol Then
                    max_vol = vol
                    vol_tic = ws.Cells(i, 1).Value
                End If
                
                'If on first row of new ticker
                If row_count = 1 Then
                    year_open = ws.Cells(i, 3).Value
                End If
            
                'Reset counters
                row_count = row_count + 1
            
            End If
  
        Next i

        'Create labels
        ws.Range("Q1").Value = "Ticker"
        ws.Range("R1").Value = "Value"
        ws.Range("P2").Value = "Max % Gain"
        ws.Range("P3").Value = "Max% Loss"
        ws.Range("P4").Value = "Max Volume"
        
        'Populate values
        ws.Range("Q2").Value = max_tic
        ws.Range("Q3").Value = min_tic
        ws.Range("Q4").Value = vol_tic
        ws.Range("R2").Value = max_val
        ws.Range("R3").Value = min_val
        ws.Range("R4").Value = max_vol
        
        'Formatting
        ws.Range("R3").NumberFormat = "0.0%"
        ws.Range("R2").NumberFormat = "0.0%"

    Next ws

End Sub
