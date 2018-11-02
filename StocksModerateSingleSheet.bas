Sub stocks():

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
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    openprice = cells(2,3).Value

    For i = 2 To lastrow

        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            Ticker = Cells(i, 1).Value
            YearChange = cells(i,6).Value - openprice 
            PercentChange = (YearChange / openprice)
            Volume = Volume + Cells(i, 7).Value

            Range("J" & Summary_Table_Row).Value = Ticker
            Range("J1").Value = "Ticker"
            Range("K" & Summary_Table_Row).Value = YearChange
            Range("K1").Value = "Yearly Change"
            Range("L" & Summary_Table_Row).Value = PercentChange
            Range("L1").Value = "Percent Change"
            Range("M" & Summary_Table_Row).Value = Volume
            Range("M1").Value = "Total Stock Volume"
            Summary_Table_Row = Summary_Table_Row + 1
            YearChange = 0
            PercentChange = 0
            Volume = 0
            openprice = cells(i+1,3).Value
        Else
            Volume = Volume + Cells(i, 7).Value
            If openprice = 0 Then
                openprice = cells(i+1,3).Value
            End If

        End If
    
    Next i

    For i = 2 to lastrow
               
        cells(i,12).style = "percent"

        If cells(i,11).value => 0 Then
            cells(i,11).interior.colorindex = 4

        Else
            cells(i,11).interior.colorindex = 3

        End If  

    Next i 
    
End Sub
