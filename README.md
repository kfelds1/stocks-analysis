# Stocks Analysis

## Background

Steve is looking to analyze the stock market but needs help finding a way to do it. Excel's Visual Basic Application is the perfect tool to analyze the performance of the 12 green stocks to show that DAQO is not the only stock worth investing in the way his parents want him to.

## Purpose

The purpose of this analysis was to help Steve show his parents different green stocks for the years 2017 and 2018 and whether or not they were worth investing in. To do this, I showed the total volume for one particular stock for either 2017 or 2018, and I also showed the return to that stock's investment using its beginning and ending prices. This return value was either positive or negative, giving a general idea on the stock's overall performance that year. After successfully completing the analysis with nested for loops, it was my job to refactor the code to find a quicker and more efficient way to achieve the same outcome.

## Results

One of the biggest differences between the two versions of code and the perhaps the driving factor behind the increase in efficiency is the creation of the arrays for 'tickerVolume,' 'tickerStartingPrices,' and 'tickerEndingPrices,' These arrays, each with a size of 12 for the 12 stocks, allowed for the storage of data for each ticker (stock). This way, we could compile data for each ticker rather than each row one by one, tremendously shortening the running time of the analysis, as seen below.

![Screen Shot 2022-10-17 at 7 00 02 PM](https://user-images.githubusercontent.com/112633146/196558271-4ca3c015-340b-48e6-ad06-d356287e8436.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/112633146/196558288-ae6572f0-fe0f-498a-8304-f178faae22e4.png)

The above and below screenshots show the time it took for the refactored code to run and its output for the years 2017 and 2018, respectively. The runtime is tremendously faster than that of the original code. 

![Screen Shot 2022-10-17 at 7 00 51 PM](https://user-images.githubusercontent.com/112633146/196558301-72ae02e3-646f-4980-ae66-ead1d82a84b0.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/112633146/196558309-3d93d145-8ac3-458f-83ff-d36c8d91cc3b.png)

### Original Code

    Sub AllStocksAnalysis()

        'Format the output sheet on the "All Stocks Analysis" worksheet

        Worksheets("All Stocks Analysis").Activate

            Range("A1").Value = "All Stocks (2018)"

            Cells(3, 1).Value = "Ticker"

            Cells(3, 2).Value = "Total Daily Volume"

            Cells(3, 3).Value = "Return"

        'Initialize an array of all tickers

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

        'Prepare for an analysis of the tickers

        'Initialize variables for starting and ending price

        Dim startingPrice As Double
        Dim endingPrice As Double

        'Activate the data worksheet

        Worksheets("2018").Activate

        'Find number of rows to loop over

        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

        'Loop through all the tickers

        For i = 0 To 11

            ticker = tickers(i)
            totalVolume = 0

            'Loop through rows in the data

            For j = 2 To RowCount

                Worksheets("2018").Activate

                'get total volume for current ticker
                If Cells(j, 1).Value = ticker Then

                    totalVolume = totalVolume + Cells(j, 8).Value

                End If

                'get starting price for current ticker

                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                    startingPrice = Cells(j, 6).Value

                End If

                'get ending price for current ticker

                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                    endingPrice = Cells(j, 6).Value

                End If

            Next j

            'Output data for current ticker

            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalVolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

        Next i

        endTime = Timer
        MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

The above code is the original code, which requires looping through the tickers and then looping through each row in order to gather the information that we need. With the refactored code, we assign the data that we need to extract (i.e. volume, starting price, and ending price) to an array. With this, we can loop through the data only once because the data for each ticker is stored in the array.

### Refactored Code

    '1a) Create a ticker Index
    'Assigning all 12 tickers to a number (i.e. 0 = AY, 1 = CSIQ, etc.)
    
    Dim tickerIndex As Integer
    tickerIndex = 0
    

    '1b) Create three output arrays
    'Each array will use the 12 ticker indexes
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single

    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    'we are calculating volume for each ticker index, so it will start at 0 and then we will add to it
    
    For j = 0 To 11
        
        tickerVolumes(j) = 0
        
    Next j
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex
        
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            'the closing price for the first row in this ticker will indicate that stock's starting price
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            'the closing price for the last row in this ticker will indicate that stock's ending price
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex. (go to next ticker in the array)
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

## Summary

### Advantages and Disadvantages of Refactoring Code

The most evident advantage to refactoring code is that it will decrease the run-time and work that a program will have to do in order to complete the analysis. The increase in efficiency is significant and is helpful when there are thousands of rows in a dataset. Refactoring code might not help just the software programming, but it also might be more concise and readable for the viewer of the code.

A disadvantage to refactoring code that comes to mind is the time and effort it might take to do so. Almost always there is more than one way to complete a task while coding, so if the route that the coder decides to go is not the most efficient, it will take more time for them to determine the most efficient way. Another disadvantage of refactoring is the risk of creating errors with already working and useful code, especially in large datasets. It might be easy to accidentally delete or mess up perfectly working code when trying to decide the most logical program.

### Advantages and Disadvantages of Refactoring the Stocks Analysis

The main advantage to refactoring this code is the decrease in run time between the original and refactored programs. Another advantage to refactoring this particular code is the elimination of the nested for loops, which made the original code a little bit difficult to read and follow. The refactored code has for loops, but there are separate from one another rather than within one another. Through this process, I feel that I have developed a deeper understanding of writing code on VBA, especially with for loops and arrays.

The disadvantage to the refactored code is the creation of another variable, 'tickerIndex.' This might pose as a source of confusion because it has a similar name to the array 'ticker.' It is important to remember that 'tickerIndex' is a variable while 'ticker' is an array, and it is helpful to think of 'tickerIndex' like the 'ticker' array but in variable form, so each position in the array is assigned a number 0-11.


