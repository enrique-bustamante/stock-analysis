# stock-analysis
Module 2 of bootcamp

### Challenge

During the course of working through the module lessons, we learned how to create nest For loops to allow for the analysis of all the stocks in the data for specific years.  While the code we wrote worked, it was inefficient due to the fact that the loop had to run 12 times over the entire data set. We knew there was a better way to approach this solution.

We were challenged with developing a new code that would run through the data sheet once, with the same yield. To do this we would have to make a few adjustments to our original code:

1. We reversed the order of the loops
2. We had to create arrays for the volume, starting price, and ending price of all stocks
3. Added for loop to sett the arrays all to zero
4. Added  for loop to display values in the appropriate cells

The new code is displayed below:


```VBA

Sub AllStockAnalysisRefactor()

'1) Format the output sheet on All Stocks Analysis worksheet

   yearValue = InputBox("What year would you like to run the analysis on?")

   Worksheets("AllStocksAnalysisRefactor").Activate
   Range("A1").Value = "All Stocks (" + yearValue + ")"

   '1a) Create a header row
   Cells(3, 1).Value = "Year"
   Cells(3, 2).Value = "Total Daily Volume"
   Cells(3, 3).Value = "Return"

   '2) Initialize array of all tickers

   Dim tickers(11) As String

   tickers(0) = "AY"
   tickers(1) = "CSIQ"
   tickers(2) = "DQ"
   tickers(3) = "ENPH"
   tickers(4) = "FSLR"
   tickers(5) = "HASI"
   tickers(6) = "JKS"
   tickers(7) = "RUN"
   tickers(8) = "SEDG"
   tickers(9) = "SPWR"
   tickers(10) = "TERP"
   tickers(11) = "VSLR"

   '3a) Initialize arrays for starting price and ending price

   Dim startingPrice(11) As Single
   Dim endingPrice(11) As Single
   Dim volume(11) As Double
   Dim tickerIndex As Integer

    'set up a variable to use in loops
   tickerIndex = 0

   '3b) Run a for loop to set all indexes to zero
   For i = 0 To 11

        startingPrice(i) = 0
        endingPrice(i) = 0
        volume(i) = 0

    Next i



   '3c) Activate data worksheet

   Worksheets(yearValue).Activate

   '3d) Get the number of rows to loop over

   RowCount = Cells(Rows.Count, "A").End(xlUp).Row

   '4) Loop through rows

          For i = 2 To RowCount

       '5) loop through indexes in the data
             For j = 0 To 11
                    tickerIndex = j
                    If Cells(i, 1).Value = tickers(tickerIndex) Then

                        volume(tickerIndex) = volume(tickerIndex) + Cells(i, 8).Value
                    End If

                    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

                        startingPrice(tickerIndex) = startingPrice(tickerIndex) + Cells(i, 6).Value

                    End If

                        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then

                            endingPrice(tickerIndex) = endingPrice(tickerIndex) + Cells(i, 6).Value

                        End If
                Next j
            Next i

   '6) Output data for display


   'Formatting
   Worksheets("AllStocksAnalysisRefactor").Activate
   Range("A3:C3").Font.Bold = True
   Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
   Range("B4:B15").NumberFormat = "#,##0"
   Range("C4:C15").NumberFormat = "0.00%"
   Columns("B").AutoFit

    Worksheets("AllStocksAnalysisRefactor").Activate

       'Loop through indexes to output the appropriate data
       For i = 0 To 11

        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = volume(i)
        Cells(4 + i, 3).Value = endingPrice(i) / startingPrice(i) - 1


         If Cells(4 + i, 3) > 0 Then

            'Change cell color to green
            Cells(4 + i, 3).Interior.Color = vbGreen

        ElseIf Cells(4 + i, 3) < 0 Then

            'Change cell color to red
            Cells(4 + i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(4 + i, 3).Interior.Color = xlNone

        End If

    Next i


End Sub
```

The result of this new code should be improved efficiency, especially when dealing with much larger datasets. As for the current data set, the only two stocks that had gains in both years were ENPH and RUN. I would recommend that Steve's parents invest in these two stocks as opposed to DQ.
