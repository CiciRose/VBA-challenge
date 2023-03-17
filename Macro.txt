Sub DogeOfWallstreet()

    For Each ws In Worksheets
        
        Dim TickerID As String
        Dim FinalRow As Long
        Dim TotalStockVolume As Double
        Dim SummaryTableRow As Long
        Dim OpenPrice As Double
        Dim ClosePrice As Double
        Dim YearlyChange As Double
        Dim PreviousAmount As Long
        Dim PercentChange As Double
        Dim GreatestPercentIncrease As Double
        Dim GreatestPercentDecrease As Double
        Dim FinalRowValue As Long
        Dim GreatestTotalVolume As Double
        
        
        TotalStockVolume = 0
        SummaryTableRow = 2
        PreviousAmount = 2
        GreatestPercentIncrease = 0
        GreatestPercentDecrease = 0
        GreatestTotalVolume = 0
        
        FinalRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        For i = 2 To FinalRow
        
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                TickerID = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryTableRow).Value = TickerID
                ws.Range("L" & SummaryTableRow).Value = TotalStockVolume
                TotalStockVolume = 0
                OpenPrice = ws.Range("C" & PreviousAmount)
                ClosePrice = ws.Range("F" & i)
                YearlyChange = ClosePrice - OpenPrice
                ws.Range("J" & SummaryTableRow).Value = YearlyChange

                If OpenPrice = 0 Then
                
                    PercentChange = 0
                    
                Else
                
                    OpenPrice = ws.Range("C" & PreviousAmount)
                    PercentChange = YearlyChange / OpenPrice
                    
                End If

                ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"
                ws.Range("K" & SummaryTableRow).Value = PercentChange

                If ws.Range("J" & SummaryTableRow).Value >= 0 Then
                
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                    
                Else
                
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                    
                End If
            
                SummaryTableRow = SummaryTableRow + 1
                PreviousAmount = i + 1
                
                End If
                
            Next i

            FinalRow = ws.Cells(Rows.Count, 11).End(xlUp).Row

            For i = 2 To FinalRow
            
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                    
                End If

                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                    
                End If

                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                    
                End If

            Next i

            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Columns("I:Q").AutoFit

    Next ws
        
End Sub
