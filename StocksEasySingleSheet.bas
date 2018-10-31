Sub stocks():

    Dim Ticker As String
    
    Dim Volume As Double
    Volume = 0
    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To lastrow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            Volume = Volume + Cells(i, 7).Value
            Range("J" & Summary_Table_Row).Value = Ticker
            Range("J1").Value = "Ticker"
            Range("K" & Summary_Table_Row).Value = Volume
            Range("K1").Value = "Total Stock Volume"
            Summary_Table_Row = Summary_Table_Row + 1
            Volume = 0
        Else
            Volume = Volume + Cells(i, 7).Value
        End If
    
    Next i
     
    
End Sub
