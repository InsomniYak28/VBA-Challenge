Sub Stocks()
‘loop through all worksheets
    For Each ws In Worksheets
        ws.Activate
‘define variables and Summary Table Row
        Dim Ticker As String
        Dim TotalStock As Long
        Dim YearChange As Double
        Dim Percentage As Double
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
‘print headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("N2").Value = "Greatest Percent Increase"
        Range("N3").Value = "Greatest Percent Decrease"
        Range("N4").Value = "Greatest Total Stock Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
‘define starting points
        TotalVolume = 0
        
        PricePointer = 2

‘find Last Row
        LR = Cells(Rows.Count, "A").End(xlUp).Row

‘create loop through rows I, J, K, L, starting after headers thru Last Row
        For i = 2 To LR
‘start with tickers: check for changes in next row
‘yearly change: subtract close price from start price columns
‘Percentage: divide yearly change by values in column C; define as percent
‘total volume: add values in column G
            If Cells(i + 1, "A").Value <> Cells(i, "A").Value Then
               Ticker = Cells(i, "A").Value
               YearlyChange = (Cells(i, "F").Value) - (Cells(PricePointer, "C").Value)
               Percentage = YearlyChange / Cells(PricePointer, "C").Value * 100
               TotalVolume = TotalVolume + Cells(i, "G").Value
‘print resulting data
               Range("I" & Summary_Table_Row).Value = Ticker
               Range("J" & Summary_Table_Row).Value = YearlyChange
               Range("K" & Summary_Table_Row).Value = "%" & Percentage
               Range("L" & Summary_Table_Row).Value = TotalVolume
‘color conditions
               If YearlyChange > 0 Then
                Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
                ElseIf YearlyChange < 0 Then
                Range("J" & Summary_Table_Row).Interior.Color = 3
                Else
                Cells(Summary_Table_Row, "J").Interior.ColorIndex = 2
                End If
            
‘repeat loop
               Summary_Table_Row = Summary_Table_Row + 1
               PricePointer = i + 1
               TotalVolume = 0
              Else
              TotalVolume = TotalVolume + Cells(i, "G").Value
              End If
             Next i
‘find Last Row again
            LR = Cells(Rows.Count, "I").End(xlUp).Row
‘loop through created data in columns I, K, and L
‘define variables
            Dim Greatest_percent_Increase
            Dim Greatest_percent_Decrease
            Dim Greatest_Total
            Dim Ticker_1
            Dim Ticker_2
            Dim Ticker_3
                
            Greatest_percent_Increase = 0
            Greatest_percent_Decrease = 0
            Greatest_Total = 0
'Greatest percent Increase w Ticker 1
            For Row = 2 To LR
                
                If Cells(Row, "K").Value > Greatest_percent_Increase Then
                    Greatest_percent_Increase = Cells(Row, "K").Value
                    Ticker_1 = Cells(Row, "I").Value
                    
                End If
            Next Row
            
            Range("O2") = Ticker_1
            Range("P2") = "%" & Greatest_percent_Increase
            
'Greatest percent Decrease w Ticker 2
            For Row = 2 To LR
                If Cells(Row, "K").Value < Greatest_percent_Decrease Then
                    Greatest_percent_Decrease = Cells(Row, "K").Value
                    Ticker_2 = Cells(Row, "I").Value
                End If
            Next Row
            
            Range("O3") = Ticker_2
            Range("P3") = "%" & Greatest_percent_Decrease
            
'Greatest Volume w Ticker 3
            For Row = 2 To LR
               If Cells(Row, "L").Value > Greatest_Total Then
                   Greatest_Total = Cells(Row, "L").Value
                   Ticker_3 = Cells(Row, "I").Value
               End If
            Next Row
            
            Range("O4") = Ticker_3
            Range("P4") = Greatest_Total
‘autofit for next worksheet
            Columns("A:P").AutoFit
        Next ws
End Sub
