# VBA-challenge
I saved my script with the export option in VBA (as VBA_Challenge.cls), it works for al ws including the Bonus challenge.

--------------------------------------------------------------------------------------------------------------
Sub ticker_ALLSHEETS()

For Each ws In Worksheets
 Dim Worksheetname As String
 Dim stock_volume As LongLong
 Dim open_value, close_value, j, i, k As LongLong


 ws.Cells(1, 9).Value = "Ticker"
 ws.Cells(1, 10).Value = "Yearly Change"
 ws.Cells(1, 11).Value = "Percent Change"
 ws.Cells(1, 12).Value = "Total Stock Volume"


 stock_volume = 0
 j = 2
 open_value = ws.Cells(2, 3)


 For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
 
    If ws.Cells(i, 1) = ws.Cells(i + 1, 1) Then
    stock_volume = stock_volume + ws.Cells(i, 7)
    

    ElseIf ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
    stock_volume = stock_volume + ws.Cells(i, 7)
    
    ws.Cells(j, 9) = ws.Cells(i, 1)
    ws.Cells(j, 12) = stock_volume
        
    ws.Cells(j, 10) = (ws.Cells(i, 6) - open_value)
    
            If open_value = 0 Then
            ws.Cells(j, 11) = 0
            Else: ws.Cells(j, 11) = (ws.Cells(i, 6) / open_value) - 1
            End If
        
            If ws.Cells(j, 10).Value < 0 Then
            ws.Cells(j, 10).Interior.ColorIndex = 3
            ws.Cells(j, 11).Interior.ColorIndex = 3
            Else: ws.Cells(j, 10).Interior.ColorIndex = 4
            ws.Cells(j, 11).Interior.ColorIndex = 4
            End If
    
              j = j + 1
              stock_volume = 0
              open_value = CLng(ws.Cells(i + 1, 3))
            
      
    End If

Next i

Next ws

For Each ws In Worksheets

ws.Cells(1, 15).Value = "Ticker"
ws.Cells(1, 16).Value = "Value"

ws.Cells(2, 14).Value = "Greatest % increase"
ws.Cells(3, 14).Value = "Greatest % decrease"
ws.Cells(4, 14).Value = "Greatest total volume"

ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(Range("K:K"))
ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(Range("K:K"))
ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(Range("L:L"))

For k = 2 To ws.Cells(Rows.Count, 7).End(xlUp).Row

If ws.Cells(k, 11).Value = ws.Cells(2, 16).Value Then
ws.Cells(2, 15).Value = ws.Cells(k, 9).Value

ElseIf ws.Cells(k, 11).Value = ws.Cells(3, 16).Value Then
ws.Cells(3, 15).Value = ws.Cells(k, 9).Value

ElseIf ws.Cells(k, 12).Value = ws.Cells(4, 16).Value Then
ws.Cells(4, 15).Value = ws.Cells(k, 9).Value

End If
Next k
Next ws


End Sub

-----------------------------------------------------------------------------------------------------------------
