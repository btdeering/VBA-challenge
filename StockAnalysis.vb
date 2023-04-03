Sub StockAnalysis()

    For Each ws In Worksheets
    
        
        Dim ticker As String
        Dim opening_price As Double
        Dim closing_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_volume As Double
        Dim last_row As Long
        Dim i As Long
        Dim j As Long
        
        
        j = 2
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        
        ticker = ws.Cells(2, 1).Value
        opening_price = ws.Cells(2, 3).Value
        total_volume = 0
        
        
        For i = 2 To last_row
        
            
            If ws.Cells(i, 1).Value <> ticker Then
            
                closing_price = ws.Cells(i - 1, 6).Value
                yearly_change = closing_price - opening_price
                If opening_price <> 0 Then
                    percent_change = yearly_change / opening_price
                Else
                    percent_change = 0
                End If
                
                ws.Cells(j, 9).Value = ticker
                ws.Cells(j, 10).Value = yearly_change
                ws.Cells(j, 11).Value = percent_change
                ws.Cells(j, 11).NumberFormat = "0.00%"
                
                ws.Cells(j, 12).Value = total_volume
                
                ticker = ws.Cells(i, 1).Value
                opening_price = ws.Cells(i, 3).Value
                total_volume = 0
                
                j = j + 1
                
            End If
        
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        Next i
        
        closing_price = ws.Cells(last_row, 6).Value
        yearly_change = closing_price - opening_price
        If opening_price <> 0 Then
            percent_change = yearly_change / opening_price
        Else
            percent_change = 0
        End If
        ws.Cells(j, 9).Value = ticker
        ws.Cells(j, 10).Value = yearly_change
        ws.Cells(j, 11).Value = percent_change
        ws.Cells(j, 12).Value = total_volume
        
        Dim max_increase As Double
        Dim max_decrease As Double
        Dim max_volume As Double
        Dim max_increase_ticker As String
        Dim max_decrease_ticker As String
        Dim max_volume_ticker As String
        
        max_increase = ws.Cells(2, 11).Value
        max_decrease = ws.Cells(2, 11).Value
        max_volume = ws.Cells(2, 12).Value
        max_increase_ticker = ws.Cells(2, 9).Value
        max_decrease_ticker = ws.Cells(2, 9).Value
        max_volume_ticker = ws.Cells(2, 9).Value
        
        For i = 2 To j - 1

    If ws.Cells(i, 11).Value > max_increase Then
        max_increase = ws.Cells(i, 11).Value
        max_increase_ticker = ws.Cells(i, 9).Value
    ElseIf ws.Cells(i, 11).Value < max_decrease Then
        max_decrease = ws.Cells(i, 11).Value
        max_decrease_ticker = ws.Cells(i, 9).Value
    End If
    
    If ws.Cells(i, 12).Value > max_volume Then
        max_volume = ws.Cells(i, 12).Value
        max_volume_ticker = ws.Cells(i, 9).Value
    End If
    
    Next i
    
    ws.Range("Q2").Value = max_increase
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").Value = max_decrease
    ws.Range("Q3").NumberFormat = "0.00%"
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = max_increase_ticker
    ws.Cells(2, 17).Value = max_increase
    
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = max_decrease_ticker
    ws.Cells(3, 17).Value = max_decrease
    
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = max_volume_ticker
    ws.Cells(4, 17).Value = max_volume

Next ws
End Sub
