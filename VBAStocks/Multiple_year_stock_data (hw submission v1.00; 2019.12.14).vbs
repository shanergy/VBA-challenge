Sub stockData()

'   /*  **************************************************  */
'       Loop through all of the sheets within the workbook
'   /*  **************************************************  */
    For Each ws In Worksheets

        '   Place column headers into sheet for the Summary Table in the columns to the right of those housing the data
        ws.Range("J1").Value = "Ticker"
        ws.Range("K1").Value = "Yearly Change"
        ws.Range("L1").Value = "Percent Change"
        ws.Range("M1").Value = "Total Stock Volume"

        '   "Starting Open Value" and "Ending Close Value" are for data validation purposes
        ws.Range("N1").Value = "Starting Open Value"
        ws.Range("O1").Value = "Ending Close Value"
        
        ws.Range("R2").Value = "Greatest % Increase"
        ws.Range("R3").Value = "Greatest % Decrease"
        ws.Range("R4").Value = "Greatest Total Volume"
        ws.Range("S1").Value = "Ticker"
        ws.Range("T1").Value = "Value"
        
        '   Create variable for the row location to place the aggregate for each stock ticker
        Dim summaryRow As Integer
        summaryRow = 2
        
        '   Create variable and define as the last row number of data to use in the for loop
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        '   Create variable to hold the stock ticker symbol
        Dim stockTicker As String

        '   Create variable to hold the change of a given stock for the given period
        Dim tickerChange As Double
        tickerChange = 0
        
        '   Create variable to hold the volume amount for the stock ticker on a day, starting at zero
        Dim stockVolume As Double
        stockVolume = 0
        
        '   Create variable to hold the stock's opening value for the start of the period
        Dim tickerStartOpenValue As Double
        tickerStartOpenValue = 0
        
        '   Create variable to hold the stock's closing value for the end of the period
        Dim tickerEndCloseValue As Double
        tickerEndCloseValue = 0
        
        '   Create variable to hold the stock ticker symbol for the stock with the greatest percent increase for the period
        Dim greatestPercentIncreaseTicker As String
        
        '   Create variable to hold the greatest percent increase value for the stock with the greatest percent increase for the period
        Dim greatestPercentIncrease As Double
        greatestPercentIncrease = 0
        
        '   Create variable to hold the stock ticker symbol for the stock with the greatest percent decrease for the period
        Dim greatestPercentDecreaseTicker As String
        
        '   Create variable to hold the greatest percent decrease value for the stock with the greatest percent decrease for the period
        Dim greatestPercentDecrease As Double
        greatestPercentDecrease = 0
        
        '   Create variable to hold the stock ticker symbol for the stock with the greatest total volume for the period
        Dim greatestTotalVolumeTicker As String
        
        '   Create variable to hold the greatest total volume for the stock with the greatest total volume for the period; set to zero
                'Dim greatestTotalVolume As Long    'Dim-ing as Long causes a run-time error; commenting out as this appears to work and run correctly
        greatestTotalVolume = 0
        
    '   /*  **************************************************  */
    '       Loop through all of rows within the current sheet
    '   /*  **************************************************  */
        For i = 2 To LastRow

            '   Lookback - grab the current stock's starting opening value (the first entry for a given stock ticker)
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                tickerStartOpenValue = ws.Cells(i, 3).Value
                '   place the starting open value of the stock ticker onto the summary table for the stock ticker
                ws.Range("N" & summaryRow).Value = tickerStartOpenValue
            End If

            '   Check to see if we are still within the same stock ticker, if it isn't...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                '   Set the Stock Ticker in the summary table
                stockTicker = ws.Cells(i, 1).Value
                
                '   Take the stock ticker and place it into the summary table
                ws.Range("J" & summaryRow).Value = stockTicker
                
                '   Add to the Stock Ticker volume
                stockVolume = stockVolume + ws.Cells(i, 7).Value
                
                '   Take the summed up volume and place it into the summary table
                ws.Range("M" & summaryRow).Value = stockVolume

                '   grab the close value of the stock ticker at the end of the stock ticker listing
                tickerEndCloseValue = ws.Cells(i, 6).Value

                '   place the last close value of the stock ticker onto the summary table for the stock ticker
                ws.Range("O" & summaryRow).Value = tickerEndCloseValue
                
                '   calculate the difference between the starting open and ending close value of the stockTicker
                tickerChange = tickerEndCloseValue - tickerStartOpenValue

                '   add tickerChange onto the summary table
                ws.Range("K" & summaryRow).Value = tickerChange

                '   conditional formatting applied; red if tickerChange < 0, green if tickerChange > 0
                '       NOTE: also decided to apply conditional formating to the "Percent Change" column in the summary table
                If tickerChange < 0 Then
                    '   set the cell color to red
                    ws.Range("K" & summaryRow).Interior.ColorIndex = 3
                    ws.Range("L" & summaryRow).Interior.ColorIndex = 3
                ElseIf tickerChange > 0 Then
                    '   set the cell color to green
                    ws.Range("K" & summaryRow).Interior.ColorIndex = 4
                    ws.Range("L" & summaryRow).Interior.ColorIndex = 4
                ElseIf tickerChange = 0 Then
                    '   set the cell color to yellow; no change
                    ws.Range("K" & summaryRow).Interior.ColorIndex = 6
                    ws.Range("L" & summaryRow).Interior.ColorIndex = 6
                End If

                '   if tickerStartOpenValue is 0, cannot be divider/denominator, results in error so use if statement to prevent this error from occuring.
                If tickerStartOpenValue <> 0 Then
                    tickerChange = (tickerEndCloseValue - tickerStartOpenValue) / tickerStartOpenValue
                    
                End If

                '   add percent change onto the summary table
                ws.Range("L" & summaryRow).Value = tickerChange
                ws.Range("L" & summaryRow).NumberFormat = "0.00%"

                '   Reset the stockVolume back to 0 to loop through next stock ticker
                stockVolume = 0

                '   Reset the tickerEndCloseValue back to 0 to loop through next stock ticker
                tickerEndCloseValue = 0
                
                '   Add onto the summary table row
                summaryRow = summaryRow + 1
            
            '   If the row immediately below the is the same stockTicker
            Else

                '   add to the stockVolume; starting value for that stock volume to sum them altogether
                stockVolume = stockVolume + ws.Cells(i, 7).Value

            End If

        Next i

        '   Loop through all of the stock ticker summaryRows
        For j = 2 To (summaryRow - 1)

            '   Look through all of the stock ticker summary values to find the greatest % increase; overwriting previous value if new value is greater than previous set value
            If ws.Cells(j, 12).Value > greatestPercentIncrease Then
                greatestPercentIncreaseTicker = ws.Cells(j, 10).Value
                greatestPercentIncrease = ws.Cells(j, 12).Value

            End If

            '   Look through all of the stock ticker summary values to find the greatest % decrease; overwriting previous value if new value is less than previous set value
            If ws.Cells(j, 12).Value < greatestPercentDecrease Then
                greatestPercentDecreaseTicker = ws.Cells(j, 10).Value
                greatestPercentDecrease = ws.Cells(j, 12).Value

            End If

            '   Look through all of the stock ticker summary values to find the greatest total volume; overwriting previous value if new value is greater than previous set value
            If ws.Cells(j, 13).Value > greatestTotalVolume Then
                greatestTotalVolumeTicker = ws.Cells(j, 10).Value
                greatestTotalVolume = ws.Cells(j, 13).Value

            End If

        Next j

        '   add greatest percent increase ticker and value to the summary table
        ws.Range("S2").Value = greatestPercentIncreaseTicker
        ws.Range("T2").Value = greatestPercentIncrease
        ws.Range("T2").NumberFormat = "0.00%"

        '   add greatest percent decrease ticker and value to the summary table
        ws.Range("S3").Value = greatestPercentDecreaseTicker
        ws.Range("T3").Value = greatestPercentDecrease
        ws.Range("T3").NumberFormat = "0.00%"

        '   add greatest total volume ticker and value to the summary table
        ws.Range("S4").Value = greatestTotalVolumeTicker
        ws.Range("T4").Value = greatestTotalVolume

    Next ws

End Sub
