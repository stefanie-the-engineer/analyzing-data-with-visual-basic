Sub MultiYearStockData()

    Dim WS As Worksheet
    
    Dim TableRow As Double
    TableRow = 2
    
    Dim RowCount As Double
    RowCount = 0
    
    Dim TickerVolume As Double
    TickerVolume = 0
    
    For Each WS In Worksheets
    
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Total Stock Volume"
    
        RowCount = WS.Cells(Rows.Count, "A").End(xlUp).Row
    
        For i = 2 To RowCount
    
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then

                ' set the ticker name
                TickerName = WS.Cells(i, 1).Value

                ' add to the ticker volume
                TickerVolume = TickerVolume + WS.Cells(i, 7).Value

                ' print the ticker name to the cell
                WS.Range("I" & TableRow).Value = TickerName

                ' print the ticker volume to the cell
                WS.Range("J" & TableRow).Value = TickerVolume
            
                ' add one to the table row
                TableRow = TableRow + 1
      
                ' reset the total ticker volume so it starts at 0 once we iterate through the next ticker name
                TickerVolume = 0

            Else

                ' iteravely sum the ticker volume
                TickerVolume = TickerVolume + WS.Cells(i, 7).Value

            End If
        
        Next i

        ' adjusts column width
        WS.Columns("I:J").AutoFit
    
    Next WS
    
End Sub
