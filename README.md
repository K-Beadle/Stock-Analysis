# Stock-Analysis

## Overview of project

  This analysis is to look at the performance of trading stock for different companies in 2017 and 2018.
  
 ### Purpose
  
  The purpose of this analysis was to find out how the company Daqos' stock performed in 2018 for a clients parents. They wanted to see the total trading volume because they believe an often traded stock accurately reflects its value. They also wanted to calculate the annual rate of return. Then, we analyzed the 11 other companies in the data set for both 2017 and 2018 so the clients parents could compare the performance of the company they were interested in to the others in the data set. 

## Results

### Analysis of Daqo's Stock Performance

![Daqo_Stock_2018](https://user-images.githubusercontent.com/78178900/111891803-9f597480-89c3-11eb-9e18-3dbcfa6fe438.png)

  The result from analyzing the Daqo stock performance for the year 2018 is that the total volume of stocks traded was 107,873,900 and if you had purchased DQ's stock at the beginning of the year and held it until the end of the year then your calculated rate of return would be -63%. From these results we can conclude that Daqo is most likely not advisable to invest your money into based on its 2018 rate of return. We can not say much about the trading volume unless we compare it to other companies.

### Analysis of All Stock Performance for 2017

![All_Stocks_2017](https://user-images.githubusercontent.com/78178900/111891566-1a219000-89c2-11eb-8605-b0a57182f59c.png)

  The result from analyzing all 12 of the stocks for the year 2017 revealed that the stock with the highest total trading volume was the company with the ticker symbol SPWR with a total trading volume of 782,187,000 and its rate of return was 23.1%; on the other hand, the trading volume for Daqos'(DQ) stock had the lowest total trading volume for the year 2017 but had the highest rate of return at 199.4%. These results portray quite the opposite representation of DQ's stock when comparing to the year 2018 that we did above this section. Out of all of the stocks analysed for this year only one had a negative rate of return at -7.2% for company TERP and the lowest positive rate of return was 5.5% for company RUN. Given DQs' volume & return, SPWR's volume & return, and TERPs' volume & return, it does not appear that a higher trading volume means a higher rate of return.
  
### Analysis of All Stock Performance for 2018

![All_Stocks_2018](https://user-images.githubusercontent.com/78178900/111891567-1e4dad80-89c2-11eb-83e7-2a472855be85.png)

  The result from analyzing all 12 of the stocks for the year 2018 revealed that the stock with the highest total trading volume was the company with the ticker symbol ENPH with a total trading volume of 607,473,500 and its rate of return was 81.9%. The lowest total trading volume was company AY at 83,079,900 with a rate of return at -7.3%. It is noteworthy that the difference in positive rates of return between 2017 and 2018 are almost opposite, there are only two companies with a positive rate of return in 2018: ENPH and RUN. All of the other companies have a negative rate of return ranging from -62.6%(DQ) to -3.5%(VSLR). So, although the previous analysis of trading for 2017 indicates that DQ was the best company to have purchased stock for, this 2018 analysis indicates it is the absolute worst to have purchased stock for. 

### Execution Time of Analysis

  To analyze this data we created a few Sub Routine's (or Macros) by using VBA in Excel. At first we created a macro to run the analysis for all the stocks which was written in a way that loops through all of the data until if finds the companies tickers symbol that we are looking for, run the analysis, store the data, and then circles back to the beginning to loop through the entire data set again and repeat the process. We also created a separate macro to format the table containing the output data so, it was a two step process to get the analyzed data and format it to easily read. The following macro was used for this original analysis process for all the stocks:
  
>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    Sub AllStocksAnalysis()
    
    Dim startTime As Single
    Dim endTime As Single
    
    'Selects the worksheet to run portion of sub routine
    Worksheets("All Stocks Analysis").Activate
    
    'Prompt to select which year to do the analysis for
    yearValue = InputBox("What year would you like to run the analysis on?")
        
    startTime = Timer
        
        'Creates a header for the worksheet
        Range("A1").Value = "All Stocks (" + yearValue + ")"
        
        'Creates header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
        
        'Creates string variable for array
        Dim tickers(12) As String
            
            'Assignes each index value in the array a ticker symbol
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
        
        'Initializing varioubles for starting and ending price
        Dim startingPrice As Double
        Dim endingPrice As Double
    
    'Selects next worksheet for next portion of sub routine
    Worksheets(yearValue).Activate
    
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
            
            'Loop through ticker symbols
            For i = 0 To 11
                ticker = tickers(i)
                totalVolume = 0
                
    'Select worksheet to loop through rows of data
    Worksheets(yearValue).Activate
                For j = 2 To RowCount
                    
                    'Find total volume for current ticker
                    If Cells(j, 1).Value = ticker Then
                        
                        'Calculates total volume
                        totalVolume = totalVolume + Cells(j, 8).Value
        
                    End If
            
                    'Find starting price for current ticker
                    If Cells(j, 1).Value = ticker And Cells(j - 1, 1) <> ticker Then
            
                        'Set starting price
                        startingPrice = Cells(j, 6).Value
            
                    End If
            
                    'Find ending price for current ticker
                    If Cells(j, 1).Value = ticker And Cells(j + 1, 1) <> ticker Then
                
                        'Set ending price
                        endingPrice = Cells(j, 6).Value
                
                    End If
                    
                Next j
    
    'Select worksheet to output data
    Worksheets("All Stocks Analysis").Activate
                'Output data for current ticker
                Cells(4 + i, 1).Value = ticker
                Cells(4 + i, 2).Value = totalVolume
                Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
            Next i
    
    endTime = Timer
    
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

The following images show how long it took to run the code for each year:

![VBA_Original_2017](https://user-images.githubusercontent.com/78178900/111893507-2e20be00-89d1-11eb-84c8-0cba5da65af9.png)
![VBA_Original_2018](https://user-images.githubusercontent.com/78178900/111893508-2fea8180-89d1-11eb-92f1-1cbb4ac65b48.png)

The concern was that if we were to use this code to analyze a larger list of stock data that it may take a long time to run because it was built to loop through the entire sheet for each ticker symbol instead of search through it once and stop at each one. Because we want to have the freedom to use this macro for larger data sets we refactored the code to make it run more efficient. The following macro contains the refactored code:

>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

    Sub AllStocksAnalysisRefactored()
    
    'Declare variables to hold sub routine timer values
    Dim startTime As Single
    Dim endTime  As Single
    
    'Prompts user to input which year to run analysis for
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    'Begins timing analysis run time
    startTime = Timer
    
    'Selects worksheet where we want to output our data table
    Worksheets("All Stocks Analysis").Activate
        
        'Creates worksheet header
        Range("A1").Value = "All Stocks (" + yearValue + ")"
    
        'Creates a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"

        'Initializes an array of all tickers
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
    
    'Activates data worksheet for year that was input by user
    Worksheets(yearValue).Activate
    
        'Gets the number of rows in the worksheet to loop over
        RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
        '1a) Create a variable to represent each index value in the array,
             'set to "0" here but will increases by "1" at the end of the loop
        Dim tickerIndex As Single

            tickerIndex = 0

        '1b) Creates variables for finding total volume and rate of return for each ticker
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
        Dim tickerEndingPrices(12) As Single
    
        '2a) Sets each ticker volume value to zero to calculate each total volume
        For i = 0 To 11

            tickerVolumes(i) = 0
    
        Next i
  
        '2b) Loops over every row in the worksheet that contains data
        For i = 2 To RowCount
    
            '3a) Increases volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
            '3b) Checks if the current row is the first row for the current tickerIndex value
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            End If
        
            '3c) Checks if the current row is the last row for the current tickerIndex value
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

                '3d If the next tickerIndex value does not match, increase the tickerIndex to loop through next ticker
                tickerIndex = tickerIndex + 1
            
            End If
    
       Next i
    
        '4) Populates cells in data table with ticker symbol, its total volume, and its rate of return
        For i = 0 To 11
        
            Worksheets("All Stocks Analysis").Activate
                Cells(4 + i, 1).Value = tickers(i)
                Cells(4 + i, 2).Value = tickerVolumes(i)
                Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
        Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
        
        'Properly formats the data so it is easier to read
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit

    'First and last row of our output data table
    dataRowStart = 4
    dataRowEnd = 15

    'Loops through data table to format cell color for "Return" column
    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            'Formats cell to green if positive number
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
            'Formats cell to red if zero or negative number
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
    
    'Stops timing analysis run time
    endTime = Timer
    
    'Creates pop up message for how long it took to run the analysis for the selected year
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

    End Sub

>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

This refactored code increased the efficiency ~tenfold shown here:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/78178900/111893991-fd428800-89d4-11eb-8f23-b31988c6edae.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/78178900/111894002-19462980-89d5-11eb-8142-67de68e994fe.png)

We created three output variables for the array for Volume and Prices and that allowed us to use the variable tickerIndex as the index value for each array. By doing that we just added 1 to its value by writing "tickerIndex = tickerIndex + 1" inside of the for loop so that it runs for each ticker symbol on every row of data only once. Instead of going through the entire sheet for every next ticker symbol like the first macro was doing. In this macro with refactored code, what was coded from 1a) to 3d) is what created more efficiency; in addition, we added the formatting macro to this macro so now it is not a two step process. If you look at the code under the 'Formatting you will see all of the code to properly format the table, this is what makes the table easy to read with proper headers and cell lines to divide the data into columns under its proper header. Refactoring the code for more efficiency was a success.

## Summary

  ### What are the advantages or disadvantages of refactoring code?
  
  ### How do these pros and cons apply to refactoring the rofinal VBA Script?
