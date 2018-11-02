Sub stocks():

    For each ws in Worksheets
    
    Dim Ticker As String
    Dim YearChange As Double
    Dim PercentChange As Double 
    Dim Volume As Double
    Dim openprice As Double

    YearChange = 0
    PercentChange = 0
    Volume = 0

    
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    openprice = ws.cells(2,3).Value

    For i = 2 To lastrow

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            YearChange = ws.cells(i,6).Value - openprice 
            PercentChange = (YearChange / openprice)
            Volume = Volume + ws.Cells(i, 7).Value

            ws.Range("J" & Summary_Table_Row).Value = Ticker
            ws.Range("J1").Value = "Ticker"
            ws.Range("K" & Summary_Table_Row).Value = YearChange
            ws.Range("K1").Value = "Yearly Change"
            ws.Range("L" & Summary_Table_Row).Value = PercentChange
            ws.Range("L1").Value = "Percent Change"
            ws.Range("M" & Summary_Table_Row).Value = Volume
            ws.Range("M1").Value = "Total Stock Volume"
            Summary_Table_Row = Summary_Table_Row + 1
            YearChange = 0
            PercentChange = 0
            Volume = 0
            openprice = ws.cells(i+1,3).Value
        Else
            Volume = Volume + ws.Cells(i, 7).Value
            If openprice = 0 Then
                openprice = ws.cells(i+1,3).Value
            End If

        End If
    
    Next i

    For i = 2 to lastrow
               
        ws.cells(i,12).style = "percent"

        If ws.cells(i,11).value => 0 Then
            ws.cells(i,11).interior.colorindex = 4

        Else
            ws.cells(i,11).interior.colorindex = 3

        End If  

    Next i 

    Next
    
End Sub
