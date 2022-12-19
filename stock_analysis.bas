Attribute VB_Name = "Module1"
Sub stock_anaylsis()
    Dim x, y, z As Double
    Dim last_row, last_ticker As Double
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        x = 1
        y = 1
        z = 2
        
        For i = 2 To last_row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1) Then
                y = y + 1
                ws.Cells(y, 9).Value = ws.Cells(i, 1).Value
                If i <> 2 Then
                    ws.Cells(y - 1, 10).Value = ws.Cells(i - 1, 6).Value - ws.Cells(z, 3).Value
                    ws.Cells(y - 1, 11).Value = ws.Cells(y - 1, 10).Value / ws.Cells(z, 3).Value
                    ws.Cells(y - 1, 12).Formula = "=SUM(" & Range(Cells(i - 1, 7), Cells(z, 7)).Address(False, False) & ")"
                    z = i
                End If
            End If
            If i = last_row Then
                ws.Cells(y, 10).Value = ws.Cells(i, 6).Value - ws.Cells(z, 3).Value
                ws.Cells(y, 11).Value = ws.Cells(y, 10).Value / ws.Cells(z, 3).Value
                ws.Cells(y, 12).Formula = "=SUM(" & Range(Cells(i, 7), Cells(z, 7)).Address(False, False) & ")"
            End If
        Next i
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        last_ticker = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K" & 2 & ":" & "K" & last_ticker))
        ws.Cells(2, 17).NumberFormat = "0.00%"
        
        ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K" & 2 & ":" & "K" & last_ticker))
        ws.Cells(3, 17).NumberFormat = "0.00%"
        
        ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L" & 2 & ":" & "L" & last_ticker))
        
        For i = 2 To last_ticker
            If ws.Cells(i, 11).Value = ws.Cells(2, 17).Value Then
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 11).Value = ws.Cells(3, 17).Value Then
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            End If
            If ws.Cells(i, 12).Value = ws.Cells(4, 17).Value Then
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
            ws.Cells(i, 11).NumberFormat = "0.00%"
            If ws.Cells(i, 10).Value = 0 Or ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(i, 10).Value < 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
            
        
        For i = 9 To 17
            ws.Columns(i).AutoFit
        Next i
    Next ws
    
End Sub

