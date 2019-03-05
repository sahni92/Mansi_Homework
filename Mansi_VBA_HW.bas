Sub StocksHW()


        
        Dim tickername As String
        Dim tickervolume As Double
        tickervolume = 0

        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Total Volume"
        
        Dim summary_ticker_row As Integer
        summary_ticker_row = 2

        

        'Count the number of rows in the first column.
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        'Looping through the rows by the ticker names
        

        For i = 2 To lastrow

            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                
                tickername = Cells(i, 1).Value

                
                tickervolume = tickervolume + Cells(i, 7).Value

                
                Range("I" & summary_ticker_row).Value = tickername

                'Print the trade volume for each ticker in the summary table
                Range("J" & summary_ticker_row).Value = tickervolume

                'Add one to the summary_ticker_row
                summary_ticker_row = summary_ticker_row + 1

                'Reset tickervolume to zero
                tickervolume = 0

            Else
              
                'Add the volume of trade
                tickervolume = tickervolume + Cells(i, 7).Value

            End If
        
        Next i

End Sub


