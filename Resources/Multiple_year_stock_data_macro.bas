Attribute VB_Name = "Module1"
Sub Stock()
    Dim ws As Worksheet
    Dim sum As Integer
    Dim year_open As Long
    Dim year_close As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_stock As Double
    Dim max_percent As Double
    Dim max_percent_ticker As String
    Dim min_percent As Double
    Dim min_percent_ticker As String
    Dim max_volume As Double
    Dim max_volume_ticker As String

    For Each ws In ThisWorkbook.Worksheets

        ws.[I1] = "Ticker"
        ws.[J1] = "Yearly Change"
        ws.[K1] = "Percent Change"
        ws.[L1] = "Total Stock Volume"
        ws.[P2] = "Greatest % Increase"
        ws.[P3] = "Greatest % Decrease"
        ws.[Q1] = "Ticker"
        ws.[R1] = "Value"
        ws.[P4] = "Max Stock Value"
        
        ws.Columns("I:R").AutoFit
        
        sum = 2
        year_open = 2
        total_stock = 0
        max_percent = 0
        min_percent = 0
        max_volume = 0
        lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        For i = 2 To lastRow
            
            total_stock = total_stock + ws.Cells(i, "G")
        
            If ws.Cells(i, "A") <> ws.Cells(i + 1, "A") Then
            
                ws.Cells(sum, "I") = ws.Cells(i, "A")
                
                yearly_change = ws.Cells(i, "F").Value - ws.Cells(year_open, "C").Value
                
                percent_change = yearly_change / (ws.Cells(year_open, "C").Value)
                
                ws.Cells(sum, "J") = yearly_change
                
                ws.Cells(sum, "K") = percent_change
                
                ws.Cells(sum, "K").NumberFormat = "0.00%"
                
                ws.Cells(sum, "L") = total_stock
                
                If (yearly_change < 0) Then
                    
                    ws.Cells(sum, "J").Interior.ColorIndex = 3
                    
                ElseIf (yearly_change > 0) Then
                    
                    ws.Cells(sum, "J").Interior.ColorIndex = 4
                    
                End If
                
                If percent_change > max_percent Then
                    
                    max_percent = percent_change
                    
                    max_percent_ticker = ws.Cells(i, "A")
                    
                End If
                
                If percent_change < min_percent Then
                    
                    min_percent = percent_change
                    
                    min_percent_ticker = ws.Cells(i, "A")
                    
                End If
                
                If total_stock > max_volume Then
                
                    max_volume = total_stock
                    
                    max_volume_ticker = ws.Cells(i, "A")
                
                End If
                
                sum = sum + 1
                
                year_open = i + 1
                
                total_stock = 0
                
            End If
            
        Next i
        
            ws.Range("Q2") = max_percent_ticker
            
            ws.Range("R2") = max_percent
            
            ws.Range("R2").NumberFormat = "0.00%"
            
            ws.Range("R3") = min_percent
            
            ws.Range("Q3") = min_percent_ticker
            
            ws.Range("R3").NumberFormat = "0.00%"
            
            ws.Range("Q4") = max_volume_ticker
            
            ws.Range("R4") = max_volume
            
    Next ws
    
End Sub

