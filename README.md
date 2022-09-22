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