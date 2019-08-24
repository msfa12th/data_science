Sub tickerVol()

    Dim ticker As String
    Dim tickerVolume As Double
    Dim lastRow As Long
    Dim numTickers As Long
    Dim tickerSummaryRow As Long
    Dim thisYear As String
    Dim i As Long
    Dim ws As Worksheet
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    
For Each ws In Worksheets
    
    numTickers = 0
    tickerVolume = 0
    tickerSummaryRow = 0
    thisYear = ws.Name 'Left(ws.Cells(2, 2).Value, 4)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    openingPrice = ws.Cells(2, 3).Value  'first opening price
    closingPrice = 0

      ' Add labels for new summary columns
      ws.Range("I1").Value = "Ticker"
      ws.Range("J1").Value = "Yearly Change"
      ws.Range("K1").Value = "Percent Change"
      ws.Range("L1").Value = "Total Stock Volume"
      
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
      
      ws.Range("P1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"

 
    
        For i = 2 To lastRow
            
            If thisYear = Left(ws.Cells(i, 2).Value, 4) Then
            
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker = ws.Cells(i, 1).Value
                    numTickers = numTickers + 1
                    tickerSummaryRow = numTickers + 1

                    tickerVolume = tickerVolume + ws.Cells(i, 7).Value
                    
                    
                    closingPrice = ws.Cells(i, 6).Value
                    yearlyChange = closingPrice - openingPrice
                    
                    If openingPrice > 0 Then
                        percentChange = yearlyChange / openingPrice
                        ws.Cells(tickerSummaryRow, 11).Value = Format(percentChange, "Percent")
                    Else
                        percentChange = 0
                        ws.Cells(tickerSummaryRow, 11).Value = "N/A"
                    End If
                    
                    ws.Cells(tickerSummaryRow, 9).Value = ticker
                    ws.Cells(tickerSummaryRow, 10).Value = yearlyChange
                    ws.Cells(tickerSummaryRow, 12).Value = tickerVolume
                    ws.Range("J:J,K:K,L:L").EntireColumn.AutoFit
                    
                     If IsEmpty(ws.Range("p2").Value) Then 'Greatest % Increase
                        ws.Range("p2").Value = ticker
                        ws.Range("q2").Value = Format(percentChange, "Percent")
                    ElseIf perchentChange > ws.Range("q2").Value Then
                        ws.Range("p2").Value = ticker
                        ws.Range("q2").Value = Format(percentChange, "Percent")
                    End If
                    
                     If IsEmpty(ws.Range("p3").Value) Then 'Greatest % Decrease
                        ws.Range("p3").Value = ticker
                        ws.Range("q3").Value = Format(percentChange, "Percent")
                    ElseIf perchentChange < ws.Range("q3").Value Then
                        ws.Range("p3").Value = ticker
                        ws.Range("q3").Value = Format(percentChange, "Percent")
                    End If
                        
                      If IsEmpty(ws.Range("p4").Value) Then 'Greatest Total Volume
                        ws.Range("p4").Value = ticker
                        ws.Range("q4").Value = tickerVolume
                    ElseIf tickerVolume > ws.Range("q4").Value Then
                        ws.Range("p4").Value = ticker
                        ws.Range("q4").Value = tickerVolume
                    End If
                        
                    ws.Range("O:O,P:P,Q:Q").EntireColumn.AutoFit
                    
                    'conditional color of cells
                    If yearlyChange > 0 Then
                        ws.Cells(tickerSummaryRow, 10).Interior.ColorIndex = 4 'green
                    ElseIf yearlyChange < 0 Then
                        ws.Cells(tickerSummaryRow, 10).Interior.ColorIndex = 3 'red
                    End If

                    openingPrice = ws.Cells(i + 1, 3).Value 'next ticker opening price
                    tickerVolume = 0 ' reset tickerVolume
                    
                Else
                    tickerVolume = tickerVolume + ws.Cells(i, 7).Value
                End If 'loop for ticker change check
            End If ' loop for year check
        Next i ' loop through each row
    
Next ws  'loop for worksheet
End Sub

