Sub stocks():

    For Each ws in Worksheets 
    
    Dim Ticker As String
    
    Dim Volume As Double
    Volume = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            Volume = Volume + ws.Cells(i, 7).Value
            ws.Range("J" & Summary_Table_Row).Value = Ticker
            ws.Range("J1").Value = "Ticker"
            ws.Range("K" & Summary_Table_Row).Value = Volume
            ws.Range("K1").Value = "Total Stock Volume"
            Summary_Table_Row = Summary_Table_Row + 1
            Volume = 0
        Else
            Volume = Volume + ws.Cells(i, 7).Value
        End If
    
    Next i

    Next
     
    
End Sub
