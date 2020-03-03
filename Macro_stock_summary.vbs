Option Explicit

Sub stock_summary()

Dim ws As Worksheet
For Each ws In Worksheets

    ws.range("I1").Value = "Ticker"
    ws.range("J1").Value = "Yearly Change"
    ws.range("K1").Value = "Percent Change"
    ws.range("L1").Value = "Total Stock Volume"
    
    Dim last_row As Long
    last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim summary_row, start_row As Long
    summary_row = 2
    start_row = 2
    
    Dim i As Long
    For i = 2 To last_row
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ws.Cells(summary_row, 9).Value = ws.Cells(i, 1).Value
    
            Dim year_closing, year_opening, year_change, percent_change As Double
            year_closing = ws.Cells(i, 6).Value
            year_opening = ws.Cells(start_row, 3).Value
    
            year_change = year_closing - year_opening
            
            If (year_opening > 0) Then
                percent_change = year_change / year_opening
            Else
                percent_change = 0
            End If
    
            Dim sum_range As String
            sum_range = "G" & start_row & ":G" & i
    
            Dim total_volume As LongLong
            total_volume = WorksheetFunction.Sum(ws.range(sum_range))
    
            ws.Cells(summary_row, 10).Value = year_change
            If ws.Cells(summary_row, 10).Value >= 0 Then
                ws.Cells(summary_row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(summary_row, 10).Interior.ColorIndex = 3
            End If
            
            ws.Cells(summary_row, 11).Value = percent_change
            ws.Cells(summary_row, 11).NumberFormat = "0.00%"
            ws.Cells(summary_row, 12).Value = total_volume
    
            start_row = i + 1
            summary_row = summary_row + 1
        End If
    Next i
    
    ws.Columns("I:L").AutoFit
    
    ws.range("O2").Value = "Greatest % Increase"
    ws.range("O3").Value = "Greatest % Decrease"
    ws.range("O4").Value = "Greatest Total Volume"
    ws.range("P1").Value = "Ticker"
    ws.range("Q1").Value = "Value"
    
    Dim total_summary_rows As Long
    total_summary_rows = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    Dim max_change, min_change As Double
    Dim max_position, min_position, max_stock_position As Long
    Dim max_volume As LongLong
    
    max_change = ws.Cells(2, 11).Value
    max_position = 2
    
    min_change = ws.Cells(2, 11).Value
    min_position = 2
    
    max_volume = ws.Cells(2, 12).Value
    max_stock_position = 2
    
    Dim j As Long
    For j = 3 To total_summary_rows
        If (ws.Cells(j, 11).Value > max_change) Then
            max_change = ws.Cells(j, 11).Value
            max_position = j
        End If
        
        If (ws.Cells(j, 11).Value < min_change) Then
            min_change = ws.Cells(j, 11).Value
            min_position = j
        End If
        
        If (ws.Cells(j, 12).Value > max_volume) Then
            max_volume = ws.Cells(j, 12).Value
            max_stock_position = j
        End If
        
    Next j
        
        ws.range("P2").Value = ws.Cells(max_position, 9).Value
        ws.range("P3").Value = ws.Cells(min_position, 9).Value
        ws.range("P4").Value = ws.Cells(max_stock_position, 9).Value
        
        ws.range("Q2").Value = max_change
        ws.range("Q2").NumberFormat = "0.00%"
        ws.range("Q3").Value = min_change
        ws.range("Q3").NumberFormat = "0.00%"
        ws.range("Q4").Value = max_volume
    
        ws.Columns("O:Q").AutoFit
        
    Next ws
    
End Sub

