Sub WallStreet_Stock_Data()


Dim ws As Worksheet

For Each ws In Worksheets

    Dim TickerSymbol As String
    Dim TotalStockVolume As Double
    Dim YearlyOpeningPrice As Double
    Dim YearlyClosingPrice As Double
    Dim summaryRow As Integer
    
    summaryRow = 2

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    Dim lastRow As Double
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        
        If (ws.Cells(i, 3).Value = 0) Then
            If (ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value) Then
                TickerSymbol = ws.Cells(i, 1).Value
            End If
        
        ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            If (ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value) Then
                YearlyOpeningPrice = ws.Cells(i, 3).Value
            End If
            
        Else
            TickerSymbol = ws.Cells(i, 1).Value
            
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            YearlyClosingPrice = ws.Cells(i, 6).Value
            
            ws.Cells(summaryRow, 9).Value = TickerSymbol
            
            ws.Cells(summaryRow, 12).Value = TotalStockVolume
            
            If (TotalStockVolume > 0) Then
                ws.Cells(summaryRow, 10).Value = YearlyClosingPrice - YearlyOpeningPrice
                
                    If (ws.Cells(summaryRow, 10).Value > 0) Then
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(summaryRow, 10).Interior.ColorIndex = 3
                    End If
                    
                ws.Cells(summaryRow, 11).Value = ws.Cells(summaryRow, 10).Value / YearlyOpeningPrice
            Else
                ws.Cells(summaryRow, 10).Value = 0
                ws.Cells(summaryRow, 11).Value = 0
            End If
            
            ws.Cells(summaryRow, 11).NumberFormat = "0.00%"
            TotalStockVolume = 0
            summaryRow = summaryRow + 1
            
        End If
    Next i

'CHALLANGES
    
    Dim GreatestTotalVolume As Double
    Dim percentincrease As Double
    Dim percentdecrease As Double
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    percentincrease = 0
    percentdecrease = 0

    For i = 2 To summaryRow
        
        If (ws.Cells(i, 11).Value > percentincrease) Then
            percentincrease = ws.Cells(i, 11).Value
            ws.Cells(2, 16) = ws.Cells(i, 9).Value
        
        ElseIf (ws.Cells(i, 11).Value < percentdecrease) Then
            percentdecrease = ws.Cells(i, 11).Value

            
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
        End If
        
    Next i
    
    ws.Cells(2, 17).Value = percentincrease
    ws.Cells(3, 17).Value = percentdecrease

    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    
    GreatestTotalVolume = 0

    summaryRow = summaryRow - 2

    For i = 2 To summaryRow
    
        If (ws.Cells(i, 12).Value > GreatestTotalVolume) Then
            GreatestTotalVolume = ws.Cells(i, 12).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
        End If
        
    Next i

    ws.Cells(4, 17).Value = GreatestTotalVolume

Next ws

End Sub

 
