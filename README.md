# Green Stocks

## Overview of Project

Our good friend Steve requested our help to analyze some stock data in order to best inform his parents on which eco-friendly investments might give the best returns. We looked at a table of data for 11 companies and 2 years which Steve had provided us, and we wrote macros in VBA which tallied the total volume of trades and final investment return for each company and either year. Then, we added buttons to make the macros easier to use, and conditional formatting to help readability. Finally, we refactored our code to make it easier for Steve to use these macros on much larger datasets.

## Results
 
### 2017 vs 2018

2017 was a good year for most of our stock returns and a great year for DQ, ENPH, FSLR, and SEDG, who all more than doubled their values. 2018 was a different story, with every stock except ENPH and RUN seeing negative returns. Analyses of both years are shown below.

![image](/Resources/2017_analysis.png)
![image](/Resources/2018_analysis.png)

### Code Performance

Our VBA scripts include a timer which measures how long the code took to run in seconds. We compared the times given by both the original scripts and the ones we refactored, and the latter were much faster.

Execution times for 2017:

Original:

![image](/Resources/VBA_Challenge_2017_unrefactored.png)

Refactored:

![image](/Resources/VBA_Challenge_2017.png)

Execution times for 2018:

Original:

![image](/Resources/VBA_Challenge_2018_unrefactored.png)

Refactored:

![image](/Resources/VBA_Challenge_2018.png)

In both cases the refactored code ran almost 10x faster. In order to achieve this, the code had to be changed so that it would wait until all the calculations had finished to read the values into cells. The original code would loop through the data until it reached a new stock ticker, at which point it would write the results of its calculation to the cells in the row of the last ticker. The main loop of this code is shown below.

```
For i = 0 To 11

    'loop through tickers

    ticker = tickers(i)
    
    totalVolume = 0
    
    Worksheets("2018").Activate
    
    For j = 2 To RowCount
    
        If Cells(j, 1).Value = ticker Then
        
            totalVolume = totalVolume + Cells(j, 8).Value
            
        End If
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            startingPrice = Cells(j, 6).Value
        
        End If
        
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
        
            endingPrice = Cells(j, 6).Value
        
        End If
    
    Next j
    
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    Cells(4 + i, 2).Value = totalVolume
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
Next i
```

After refactoring, the code initialized arrays with a dimension of 12, the number of stock tickers, before everything else. This allowed the results for each ticker to be added to the arrays within the loop, and then those results were read out to cells at the very end. Some of the refactored code is shown below.

```
For i = 2 To RowCount
        ticker = tickers(tickerIndex)
    
        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        If Cells(i + 1, 1).Value <> ticker And Cells(i, 1).Value = ticker Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
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
```