# Analyzing Stock Data with VBA

## Overview of Project

This project involved the use of Visual Basic for Applications (VBA), a programming language for Microsoft Office applications. VBA programming language is used to automate various tasks to create time efficiencies and reduce accidents and errors. For this work, VBA scripts were used with Microsoft Excel to read and write to cells and worksheets, make calculations, and use logic to perform analyses across multiple stocks from 2017 and 2018. VBA macros were created that trigger pop-ups and inputs, read and change cell values, and format cells. Various foundational coding skills were used, including the use of for loops, nested for loops, and conditionals to direct logic flow, as well as syntax recollection, problem decomposition, debugging, and refactoring. 

### Purpose

The purpose of this project is to analyze the performance DAQO New Energy Corp stock as well as various other green energy stocks, including those in the hydro, geothermal, wind, and bio space, in order to guide future investment opportunities. Following the creation of macros to automate the calculation, formatting, and analysis of stock data, the code was refactored to improve efficiency and reduce run times.  

## Results

Following the refactoring of the script, execution times were significantly trimmed as seen in the image below. Code now executes roughly 12x faster. 

![Refactored_Improvement.png](https://github.com/tysonseang/stock-analysis/blob/main/Resources/Refactored_Improvement.png)

The formatted 2017 and 2018 analysis can at the file titled VBA_Challenege located [here](https://github.com/tysonseang/stock-analysis/blob/main/VBA_Challenge.xlsm). Pre- and post-refactoring code is shown below to showcase the updates. 

### Code Prior to Refactoring

'''

    Sub AllStocksAnalysis()

        Dim startTime As Single
        Dim endTime As Single

         'Create input box for year

         yearValue = InputBox("What year would you like to run the analysis on?")

        'set startTime variable equal to the Timer function in order to start the clock

        startTime = Timer

        '1) Format the output sheet on the "All Stocks Analysis" Worksheet

            Worksheets("All Stocks Analysis").Activate

        Range("A1").Value = "All Stocks (" + yearValue + ")"

            Cells(3, 1).Value = "Ticker"
            Cells(3, 2).Value = "Total Daily Volume"
            Cells(3, 3).Value = "Return"

        '2) Initialize an array of all tickers

            Dim tickers(12) As String

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

        '3a) Initialize variables for the starting price and ending price

            Dim startingPrice As Double
            Dim endingPrice As Double

        '3b) Activate the data worksheet

            Worksheets(yearValue).Activate

        '3c) Find the number of rows to loop over

            RowCount = Cells(Rows.Count, "A").End(xlUp).Row

        '4) Loop through the tickers

            For i = 0 To 11
                ticker = tickers(i)
                totalVolume = 0

            '5) Loop through the rows in the data

            Worksheets(yearValue).Activate
                For j = 2 To RowCount

            '5a) Find total volume for the current ticker

            If Cells(j, 1).Value = ticker Then

                totalVolume = totalVolume + Cells(j, 8).Value

                End If

            '5b) Find starting price for the current ticker

            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1) = ticker Then

                startingPrice = Cells(j, 6).Value

                End If

            '5c) Find ending price for the current ticker

            If Cells(j + 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                endingPrice = Cells(j, 6).Value

                End If

            Next j

        '6) Out the data for the current ticker

            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

        Next i

             endTime = Timer
            MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub   
'''

### Code After Refactoring

'''

    Sub AllStocksAnalysisRefactored()
        Dim startTime As Single
        Dim endTime  As Single

        yearValue = InputBox("What year would you like to run the analysis on?")

        startTime = Timer

        'Format the output sheet on All Stocks Analysis worksheet
        Worksheets("All Stocks Analysis").Activate

        Range("A1").Value = "All Stocks (" + yearValue + ")"

        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

        'Initialize array of all tickers
        Dim tickers(12) As String

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

        'Activate data worksheet
        Worksheets(yearValue).Activate

        'Get the number of rows to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

        '1a) Create a ticker Index
        Dim tickerIndex As Integer
        tickerIndex = 0
        ticker = tickers(tickerIndex)


        '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single

        ''2a) Create a for loop to initialize the tickerVolumes to zero.
          tickerVolumes(tickerIndex) = 0

        ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount

            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value


            '3b) Check if the current row is the first row with the selected tickerIndex.
            'If  Then

                If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
                    tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

                End If

            '3c) check if the current row is the last row with the selected ticker
             'If the next row’s ticker doesn’t match, increase the tickerIndex.
            'If  Then

                If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
                    tickerEndingPrices(tickerIndex) = Cells(i, 6).Value


                '3d Increase the tickerIndex.
                    tickerIndex = tickerIndex + 1

                End If

        Next i

        '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
        For i = 0 To 11

            Worksheets("All Stocks Analysis").Activate

            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

        Next i

        'Formatting
        Worksheets("All Stocks Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit

        dataRowStart = 4
        dataRowEnd = 15

        For i = dataRowStart To dataRowEnd

            If Cells(i, 3) > 0 Then

                Cells(i, 3).Interior.Color = vbGreen

            Else


                Cells(i, 3).Interior.Color = vbRed

            End If

        Next i

        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub
'''

## Summary
Refactoring code can boost system performance and reduce the time needed to execute large VBA macros. This becomes increasingly important with larger datasets and datasets that are constantly updating. Due to the additional investments of time and work required to refactor code, it is important to weigh the pros and cons. For this project, refactored code ran roughly 12x faster than the original VBA script. This could prove beneficial if the code were used on larger datasets. However, refactoring was a time-intensive process, which was further impacted by my lack of significant experience writing VBA scripts.
