Sub Stock_Market_Analysis()

Dim lastRowStockData As Long
Dim opening_price, closing_price, yearly_change As Double
Dim greatest_per_increase, greatest_per_decrease  As Double
Dim ticker_per_increase, ticker_per_decrease, ticker_total_volume As String
Dim great_total_volume, StockTotal  As LongLong
Dim percent_change As String

'Loop through every worksheet and process the state contents.
   For Each ws In Worksheets
       ' Find the last row of each worksheet and Subtract one to return the number of rows without header
       lastRowStockData = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
       greatest_per_increase = 0
       greatest_per_decrease = 0



        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

       'Intializing the value of StockTotal variable
       StockTotal = ws.Cells(2, 7)
       great_total_volume = 0
       a = 2
       j = 2
       For i = 2 To lastRowStockData
          ws.Cells(a, 9).Value = ws.Cells(j, 1).Value ' This is to populate Ticker
          opening_price = ws.Cells(j, 3).Value

            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                    closing_price = ws.Cells(i, 6).Value 'closing price
                    yearly_change = closing_price - opening_price 'Yearly Change
                    ws.Cells(a, 10).Value = yearly_change
                    ws.Cells(a, 10).NumberFormat = "0.00000000"
                    If ws.Cells(a, 10).Value < 0 Then
                        'Assign red color
                        ws.Cells(a, 10).Interior.ColorIndex = 3
                    Else
                        'assign Green color to cell
                        ws.Cells(a, 10).Interior.ColorIndex = 4
                    End If

                    If (opening_price <> 0) Then
                    percent_change = Format((closing_price - opening_price) / opening_price, "Percent") ' percent difference
                    ws.Cells(a, 11).Value = percent_change
                        If (greatest_per_increase < ((closing_price - opening_price) / opening_price)) Then
                        greatest_per_increase = ((closing_price - opening_price) / opening_price)
                        ticker_per_increase = ws.Cells(a, 9).Value
                        End If

                        If (greatest_per_decrease > ((closing_price - opening_price) / opening_price)) Then
                        greatest_per_decrease = ((closing_price - opening_price) / opening_price)
                        ticker_per_decrease = ws.Cells(a, 9).Value
                        End If
                    Else
                        percent_change = 0
                    End If

            ws.Cells(a, 12).Value = StockTotal
            StockTotal = ws.Cells(i+1, 7).Value
            a = a + 1
            j = i + 1
            Else
                StockTotal = StockTotal + ws.Cells(i + 1, 7).Value
            End If

            If (great_total_volume < StockTotal) Then
                   great_total_volume = StockTotal
                   ticker_total_volume = ws.Cells(a, 9).Value
            End If
        Next i

        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(2, 16).Value = ticker_per_increase
        ws.Cells(2, 17).Value = Format(greatest_per_increase, "Percent")

        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = ticker_per_decrease
        ws.Cells(3, 17).Value = Format(greatest_per_decrease, "Percent")

        ws.Cells(4, 15).Value = "Greatest total volume"
        ws.Cells(4, 16).Value = ticker_total_volume
        ws.Cells(4, 17).Value = great_total_volume

     ws.Columns("A:Q").AutoFit
   Next ws
   MsgBox ("Stock Market Analysis Processing Complete !!")
End Sub
