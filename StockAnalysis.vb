Sub StockAnalysis():
For Each ws In ThisWorkbook.Worksheets
  'Declaring the Variables
  Dim Ticker As String
  Dim YearlyChange As Double
  Dim PercentChange As Double
  Dim TotalVolume As Double
  Dim LastRow As Long
  Dim SummaryRow As Long
  Dim OpeningPrice As Double
  Dim ClosingPrice As Double
  Dim MaxIncrease As Double
  Dim MaxDecrease As Double
  Dim MaxVolume As Double
  Dim MaxIncreaseTicker As String
  Dim MaxDecreaseTicker As String
  Dim MaxVolumeTicker As String
  Dim OpenRow As Double
 
'naming the cells
ws.Range("i1").Value = "Ticker"
ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total Stock Volume"

'finding lastrow and assigning where the data starts from

LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
OpenRow = 2
SummaryRow = 2

'getting the ticker, open price, closingprice and yearly change

 For i = 2 To LastRow
 
     TotalVolume = TotalVolume + ws.Cells(i, 7).Value

     If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     Ticker = ws.Cells(i, 1).Value
     OpeningPrice = ws.Cells(OpenRow, 3).Value
     ClosingPrice = ws.Cells(i, 6).Value
     YearlyChange = ClosingPrice - OpeningPrice
 OpenRow = i + 1
    If OpeningPrice <> 0 Then
         PercentChange = (YearlyChange / OpeningPrice)
    Else
          PercentChange = 0
 
    End If

 ws.Cells(SummaryRow, 9).Value = Ticker
 ws.Cells(SummaryRow, 10).Value = YearlyChange
 ws.Cells(SummaryRow, 11).Value = PercentChange
 ws.Cells(SummaryRow, 12).Value = TotalVolume
 ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
 
 TotalVolume = 0
 
 If YearlyChange > 0 Then
   ws.Cells(SummaryRow, 10).Interior.Color = RGB(0, 255, 0)
 ElseIf YearlyChange < 0 Then
    ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
 End If
             
 'Update variables for the next ticker
 SummaryRow = SummaryRow + 1
 
 
        End If
 Next i
       
    ' Find the greatest % increase, % decrease, and total volume
        MaxIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & SummaryRow))
        MaxDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & SummaryRow))
        MaxVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & SummaryRow))
       
        ' Find the Tickers associated with the max values
        MaxIncreaseTicker = ws.Cells(Application.WorksheetFunction.Match(MaxIncrease, ws.Range("K2:K" & SummaryRow), 0) + 1, 9).Value
        MaxDecreaseTicker = ws.Cells(Application.WorksheetFunction.Match(MaxDecrease, ws.Range("K2:K" & SummaryRow), 0) + 1, 9).Value
        MaxVolumeTicker = ws.Cells(Application.WorksheetFunction.Match(MaxVolume, ws.Range("L2:L" & SummaryRow), 0) + 1, 9).Value
       
        ' Output the greatest % increase, % decrease, and total volume
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(2, 16).Value = MaxIncreaseTicker
        ws.Cells(3, 16).Value = MaxDecreaseTicker
        ws.Cells(4, 16).Value = MaxVolumeTicker
        ws.Cells(2, 17).Value = MaxIncrease
        ws.Cells(3, 17).Value = MaxDecrease
        ws.Cells(4, 17).Value = MaxVolume
        ws.Cells(2, 17).NumberFormat = "0.00%"
        ws.Cells(3, 17).NumberFormat = "0.00%"

       
       Next ws
 
End Sub

