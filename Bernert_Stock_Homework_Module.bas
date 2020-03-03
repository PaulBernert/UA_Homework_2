Attribute VB_Name = "Module1"
Sub Parser_()

'MUST ITERATE OVER ALL WORKSHEETS
Dim ws As Worksheet

For Each ws In Worksheets

    'Assign ticker symbol as string
    Dim ticker_symbol As String
    'Assign last row and summary_row as integer
    Dim last_row, summary_row As Integer
    'Assign open and close values as double
    Dim year_open, year_close As Double
    'Assign vol as float/long
    Dim vol As Variant

    'Assign values
    current_row = 0
    summary_row = 2
    vol = 0
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Add labels and misc. text
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Year Open"
    ws.Range("K1").Value = "Year Close"
    ws.Range("N1").Value = "Volume"
    ws.Range("L1").Value = "Annual $ Chng"
    ws.Range("M1").Value = "Annual % Chng"
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    
    'Formatting
    
    
    'Loop over rows
    For i = 2 To last_row

        'Check if values match, if not...
        If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
        
            'Assign values
            vol = 0
            ticker_symbol = ws.Cells(i, 1).Value
            year_open = ws.Cells(i - current_row, 3).Value
            
            'Write the brand, open and close values
            ws.Range("I" & summary_row).Value = ticker_symbol
            ws.Range("J" & summary_row).Value = year_open
            
            'Add one to the summary row, reset current row
            summary_row = summary_row + 1
            current_row = 0
            
        'If they match..
        ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
            
            'Assign values
            year_close = ws.Cells(i + 1, 6).Value
            vol = vol + ws.Cells(i + 1, 7).Value
            
            'Write the year close value
            ws.Range("K" & summary_row).Value = year_close
            ws.Range("N" & summary_row).Value = vol
            
            'Add one to the current row
            current_row = current_row + 1

        End If
        
    Next i
    
    'Assign necessary items
    Dim open_val, close_val As Double
    Dim soln_sub, soln_pct As Double
    Dim short_row, sum_row As Integer
    
    'Assign values
    short_row = ws.Range("I1", ws.Range("I1").End(xlDown)).Rows.Count
    
    For j = 2 To short_row
    
        'Assign values
        open_val = ws.Cells(j, 10).Value
        close_val = ws.Cells(j, 11).Value
        
        'Calculate values
        If (open_val <> 0 And close_val <> 0) Then
            soln_sub = (close_val - open_val)
            soln_pct = (close_val / open_val) - 1
        Else
            soln_sub = 0
            soln_pct = 0
        End If
        
        'Populate values
        ws.Cells(j, 12).Value = soln_sub
        ws.Cells(j, 13).Value = soln_pct
        
        'Stylistic changes
        If (soln_sub < 0) Then
            ws.Cells(j, 12).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 12).Interior.ColorIndex = 10
        End If
        
        If (soln_pct < 0) Then
            ws.Cells(j, 13).Interior.ColorIndex = 3
        Else
            ws.Cells(j, 13).Interior.ColorIndex = 10
        End If
        
    Next j
    
    'Assign necessary items
    Dim max_increase_val_sheet, max_decrease_val_sheet As Double
    Dim max_increase_val_all, max_decrease_val_all As Double
    Dim max_vol_sheet As Variant
    Dim max_vol_all As Variant
    Dim max_increase_ticker_sheet, max_decrease_ticker_sheet, max_vol_ticker_sheet As String
    Dim max_increase_ticker_all, max_decrease_ticker_all, max_vol_ticker_all As String
    
    'Default values
    max_increase_val_sheet = 0
    max_decrease_val_sheet = 0
    max_vol_val_sheet = 0
        
    'Check for max positive change
    For k = 2 To short_row
        
        If (ws.Cells(k + 1, 13).Value > max_increase_val_sheet) Then
            max_increase_val_sheet = ws.Cells(k + 1, 13).Value
            max_increase_ticker_sheet = ws.Cells(k + 1, 9).Value
            If (max_increase_val_sheet > max_increase_val_all) Then
                max_increase_val_all = max_increase_val_sheet
                max_increase_ticker_all = max_increase_ticker_sheet
            End If
        End If
        
    Next k
        
    'Check for max negative change
    For l = 2 To short_row
        
        If (ws.Cells(l + 1, 13).Value < max_decrease_val_sheet) Then
            max_decrease_val_sheet = ws.Cells(l + 1, 13).Value
            max_decrease_ticker_sheet = ws.Cells(l + 1, 9).Value
            If (max_decrease_val_sheet > max_decrease_val_all) Then
                max_decrease_val_all = max_decrease_val_sheet
                max_decrease_ticker_all = max_decrease_ticker_sheet
            End If
        End If
            
    Next l
        
    'Check max volume
    For m = 2 To short_row
        
        If (ws.Cells(m + 1, 14).Value > max_vol_val_sheet) Then
            max_vol_val_sheet = ws.Cells(m + 1, 14).Value
            max_vol_ticker_sheet = ws.Cells(m + 1, 9).Value
            If (max_vol_val_sheet > max_vol_val_all) Then
                max_vol_val_all = max_vol_val_sheet
                max_vol_ticker_all = max_vol_ticker_sheet
            End If
        End If
            
    Next m

    'Populate values
    ws.Cells(2, 17).Value = max_increase_ticker_sheet
    ws.Cells(2, 18).Value = max_increase_val_sheet
    ws.Cells(3, 17).Value = max_decrease_ticker_sheet
    ws.Cells(3, 18).Value = max_decrease_val_sheet
    ws.Cells(4, 17).Value = max_vol_ticker_sheet
    ws.Cells(4, 18).Value = max_vol_val_sheet

Next ws

End Sub
